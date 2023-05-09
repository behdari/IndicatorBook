using IndicatorBook.Classes;
using IndicatorBook.Resources;
using IndicatorBook.Sent;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using SaveOptions = System.Xml.Linq.SaveOptions;

namespace IndicatorBook.Recived
{
    public partial class RecivedForm : Form
    {
        float FontSize = 10.0f;
        private bool isInit = true;
        public string _fileSharingBase;

        public ListViewColumnSorter LvwColumnSorter { get; private set; }

        public RecivedForm()
        {
            InitializeComponent();
            //listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            InitializeList();
            if (Directory.Exists("D:\\file sharing"))
            {
                _fileSharingBase = "D:\\file sharing";
            }
            else
            {
                var basePath = Directory.GetCurrentDirectory();
                _fileSharingBase = Directory.GetParent(basePath).Parent.ToString();
            }
        }

        private void InitializeList()
        {
            listView1.MultiSelect = false;
            listView1.GridLines = true;
            listView1.AllowColumnReorder = true;
            listView1.LabelEdit = true;
            listView1.FullRowSelect = true;
            listView1.Sorting = SortOrder.Ascending;
            listView1.View = View.Details;
            LvwColumnSorter = new ListViewColumnSorter();
            listView1.ListViewItemSorter = LvwColumnSorter;
        }

        #region Buttons

        void buttonNew_Click(object sender, EventArgs e)
        {
            try
            {
                int previousRowNumber = 0;

                if (listView1.Items.Count > 0)
                {
                    previousRowNumber = Convert.ToInt32(listView1.Items[0].SubItems[7].Text);
                }

                RecivedItemForm newForm = new RecivedItemForm(true, false, _fileSharingBase);
                newForm.Font = new Font(this.Font.Name, this.FontSize, this.Font.Style, this.Font.Unit, this.Font.GdiCharSet, this.Font.GdiVerticalFont);
                newForm.Text = "افزودن آیتم جدید";
                //newForm.lableRegDate.Text = christianToolStripMenuItem.Checked ? DateTime.Now.ToString() : ConvertToPersianDate(DateTime.Now.ToString());
                newForm.textBoxRowNumber.Text = (previousRowNumber + 1).ToString();
                newForm.textBoxPreviousRowNumber.Text = previousRowNumber.ToString();

                newForm.persianDatePickerDate.GeorgianDate = DateTime.Today;
                newForm.ShowDialog();
                CheckTextBoxSearchAndFillListView();
                //int contactsNumbers = Variables.xDocument.Descendants("Item").Where(q => q.Attribute("UserID").Value == Variables.CurrentUserID).Count();
                //this.Text = Variables.Caption + Variables.CurrentUserName + " : " + contactsNumbers.ToString() + " Contacts";
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                Messages.error(ref file_info, ex.Message, this);
            }
        }

        void buttonClearSearchTextBox_Click(object sender, EventArgs e)
        {
            textBoxSearch.Text = "";
            CheckTextBoxSearchAndFillListView();
        }

        void buttonEdit_Click(object sender, EventArgs e)
        {
            try
            {
                if (listView1.SelectedItems.Count < 1) return;

                string id = listView1.SelectedItems[0].Name;

                //var item = (from q in Variables.xDocument.Descendants("Item")
                //            where q.Attribute("UserID").Value == Variables.CurrentUserID && q.Attribute("ID").Value == id
                //            select q).First();
                //if (item == null) return;

                RecivedItemForm editForm = new RecivedItemForm(false, true, _fileSharingBase);

                editForm.Font = new Font(this.Font.Name, this.FontSize, this.Font.Style, this.Font.Unit, this.Font.GdiCharSet, this.Font.GdiVerticalFont);
                editForm.Text = "ویرایش آیتم";

                editForm.textBoxRowNumber.Text = listView1.SelectedItems[0].SubItems[7].Text;
                editForm.textBoxPreviousRowNumber.Text = listView1.SelectedItems[0].SubItems[6].Text;
                //editForm.textBoxDate.Text = listView1.SelectedItems[0].SubItems[5].Text;
                var splitedDate = listView1.SelectedItems[0].SubItems[5].Text.Split('/');
                editForm.persianDatePickerDate.PersianDate = new PersianDateInfo(Convert.ToInt32(splitedDate[0].Trim()), // year
                                                                                 Convert.ToInt32(splitedDate[1].Trim()), // month
                                                                                 Convert.ToInt32(splitedDate[2].Trim())); // day
                editForm.textBoxLetterOwners.Text = listView1.SelectedItems[0].SubItems[4].Text;
                editForm.textBoxDescription.Text = listView1.SelectedItems[0].SubItems[3].Text;

                var hasAttachment = listView1.SelectedItems[0].SubItems[2].Text == "دارد";
                if (hasAttachment)
                {
                    editForm.radioButtonHasAttachment.Checked = true;
                }
                else
                {
                    editForm.radioButtonHasNotAttachment.Checked = true;
                }

                editForm.textBoxRecivedLetterNumber.Text = listView1.SelectedItems[0].SubItems[1].Text;
                var splitedRecivedLetterDate = listView1.SelectedItems[0].SubItems[0].Text.Split('/');
                editForm.persianDatePickerRecivedLetterDate.PersianDate = new PersianDateInfo(Convert.ToInt32(splitedRecivedLetterDate[0].Trim()), // year
                                                                                              Convert.ToInt32(splitedRecivedLetterDate[1].Trim()), // month
                                                                                              Convert.ToInt32(splitedRecivedLetterDate[2].Trim())); // day

                editForm.ItemID = id;

                editForm.ShowDialog();

                CheckTextBoxSearchAndFillListView();
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                Messages.error(ref file_info, ex.Message, this);
            }
        }

        void buttonDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (listView1.SelectedItems.Count < 1) return;
                if (MessageBox.Show("آیا مطمئن هستید که میخواهید این مورد را حذف کنید؟", "هشدار", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No) return;
                var id = Convert.ToInt32(listView1.SelectedItems[0].Name);

                RecivedRepository.Instance.Delete(id);

                CheckTextBoxSearchAndFillListView();
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                Messages.error(ref file_info, ex.Message, this);
            }
        }

        #endregion

        #region Menu Strip Events

        #region Settings

        void rightToLeftToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                rightToLeftToolStripMenuItem.Checked = true;
                leftToRightToolStripMenuItem.Checked = false;
                textBoxSearch.RightToLeft = RightToLeft.Yes;
                listView1.RightToLeft = RightToLeft.Yes;

                var query = (from q in Variables.xDocument.Descendants("Setting")
                             where q.Attribute("UserID").Value == Variables.CurrentUserID
                             select q).First();
                query.Attribute("RightToLeft").Value = "Yes";
                TripleDES.EncryptToFile(Variables.xDocument.ToString(SaveOptions.DisableFormatting), Variables.DBFile, TripleDES.ByteKey, TripleDES.IV);
                //Variables.xDocument.Save("debug.xml");
            }
            catch { }
        }

        void leftToRightToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                leftToRightToolStripMenuItem.Checked = true;
                rightToLeftToolStripMenuItem.Checked = false;
                textBoxSearch.RightToLeft = RightToLeft.No;
                listView1.RightToLeft = RightToLeft.No;

                var query = (from q in Variables.xDocument.Descendants("Setting")
                             where q.Attribute("UserID").Value == Variables.CurrentUserID
                             select q).First();
                query.Attribute("RightToLeft").Value = "NO";
                TripleDES.EncryptToFile(Variables.xDocument.ToString(SaveOptions.DisableFormatting), Variables.DBFile, TripleDES.ByteKey, TripleDES.IV);
            }
            catch { }
        }

        void toolStripMenuItemFontSize_Click(object sender, EventArgs e)
        {
            try
            {
                toolStripMenuItemFontSize8.Checked = toolStripMenuItemFontSize10.Checked = toolStripMenuItemFontSize12.Checked = toolStripMenuItemFontSize14.Checked = toolStripMenuItemFontSize16.Checked = toolStripMenuItemFontSize18.Checked = false;
                ToolStripMenuItem menuItem = sender as ToolStripMenuItem;
                menuItem.Checked = true;
                this.FontSize = float.Parse(menuItem.Text.Trim());
                if (this.Font.Size != this.FontSize)
                {
                    this.Font = new Font(this.Font.Name, this.FontSize, this.Font.Style, this.Font.Unit, this.Font.GdiCharSet, this.Font.GdiVerticalFont);
                    var query = (from q in Variables.xDocument.Descendants("Setting")
                                 where q.Attribute("UserID").Value == Variables.CurrentUserID
                                 select q).First();
                    query.Attribute("FontSize").Value = this.FontSize.ToString();
                    TripleDES.EncryptToFile(Variables.xDocument.ToString(SaveOptions.DisableFormatting), Variables.DBFile, TripleDES.ByteKey, TripleDES.IV);
                    //Variables.xDocument.Save("debug.xml");
                }
            }
            catch { }
        }

        void christianToolStripMenuItem_Click(object sender, EventArgs e)
        {
            christianToolStripMenuItem.Checked = true;
            persianToolStripMenuItem.Checked = false;

            var query = (from q in Variables.xDocument.Descendants("Setting")
                         where q.Attribute("UserID").Value == Variables.CurrentUserID
                         select q).First();
            query.Attribute("Dates").Value = "Christian";
            TripleDES.EncryptToFile(Variables.xDocument.ToString(SaveOptions.DisableFormatting), Variables.DBFile, TripleDES.ByteKey, TripleDES.IV);
            //Variables.xDocument.Save("debug.xml");
        }

        void persianToolStripMenuItem_Click(object sender, EventArgs e)
        {
            christianToolStripMenuItem.Checked = false;
            persianToolStripMenuItem.Checked = true;

            var query = (from q in Variables.xDocument.Descendants("Setting")
                         where q.Attribute("UserID").Value == Variables.CurrentUserID
                         select q).First();
            query.Attribute("Dates").Value = "Persian";
            TripleDES.EncryptToFile(Variables.xDocument.ToString(SaveOptions.DisableFormatting), Variables.DBFile, TripleDES.ByteKey, TripleDES.IV);
            //Variables.xDocument.Save("debug.xml");
        }

        #endregion

        void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        void newUserToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var sentForm = SwitchToSent();

            sentForm.newUserToolStripMenuItem_Click(null, null);
        }

        void changeUserToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SwitchToSent(true);
        }

        void changeInfoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                UserForm changeInfoForm = new UserForm(false, false, true, false);
                changeInfoForm.Font = new Font(this.Font.Name, this.FontSize, this.Font.Style, this.Font.Unit, this.Font.GdiCharSet, this.Font.GdiVerticalFont);

                var userElement = from q in Variables.xDocument.Descendants("User")
                                  where q.Attribute("ID").Value == Variables.CurrentUserID
                                  select q;
                string username = userElement.First().Attribute("UserName").Value;
                string email = userElement.First().Attribute("Email").Value;

                changeInfoForm.textBoxUsername.Text = username;
                changeInfoForm.textBoxEmail.Text = email;
                changeInfoForm.ShowDialog();

                if (Variables.CurrentUserName != "" && Variables.CurrentUserID != "")
                {
                    int contactsNumbers = Variables.xDocument.Descendants("Item").Where(q => q.Attribute("UserID").Value == Variables.CurrentUserID).Count();
                    this.Text = Variables.Caption + Variables.CurrentUserName + " : " + contactsNumbers.ToString() + " Contacts";
                    DisableEnableControls(true);
                }
                else
                    DisableEnableControls(false);
            }
            catch (Exception ex)
            {
                DisableEnableControls(false);
                StackFrame file_info = new StackFrame(true);
                Messages.error(ref file_info, ex.Message, this);
            }
        }

        void aboutProgrammerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.mds-soft.persianblog.ir/");
        }

        #endregion

        private void FillListView(List<RecivedLetter> items)
        {
            foreach (var item in items)
            {
                ListViewItem listViewItems;

                //listViewItems = new ListViewItem(new string[]
                //                { item.RowNumber.ToString(),
                //                  item.PreviousRowNumber.ToString(),
                //                  item.Date.ToString(),
                //                  item.LetterOwners.ToString(),
                //                  item.Description,
                //                  item.HasAttachment.ToString(),
                //                  item.RecivedLetterNumber.ToString(),
                //                  item.RecivedLetterDate.ToString() });

                listViewItems = new ListViewItem(new string[]
                                { item.RecivedLetterDate.ToString(),
                                  item.RecivedLetterNumber.ToString(),
                                  item.HasAttachment? "دارد" : "ندارد",
                                  item.Description,
                                  item.LetterOwners.ToString(),
                                  item.Date.ToString(),
                                  item.PreviousRowNumber.ToString(),
                                  item.RowNumber.ToString() });

                listViewItems.Name = item.Id.ToString();

                listView1.Items.Add(listViewItems);
            }

            SetColorOfTable();
        }

        private void SetColorOfTable()
        {
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                if (i % 2 == 1)
                    listView1.Items[i].BackColor = Color.White;

                else
                    listView1.Items[i].BackColor = Color.FromArgb(255, 204, 255);
            }
        }

        void ApplySettings()
        {
            try
            {
                if (Variables.xDocument == null)
                {
                    DisableEnableControls(false);
                    return;
                }

                var settings = (from q in Variables.xDocument.Descendants("Setting")
                                where q.Attribute("UserID").Value == Variables.CurrentUserID
                                select q).FirstOrDefault();

                if (settings == null)
                {
                    return;
                }

                if (settings.Attribute("RightToLeft").Value == "Yes")
                    rightToLeftToolStripMenuItem_Click(null, null);
                else
                    leftToRightToolStripMenuItem_Click(null, null);

                if (settings.Attribute("Dates").Value == "Persian")
                {
                    persianToolStripMenuItem.Checked = true;
                    christianToolStripMenuItem.Checked = false;
                }
                else
                {
                    persianToolStripMenuItem.Checked = false;
                    christianToolStripMenuItem.Checked = true;
                }

                this.FontSize = float.Parse(settings.Attribute("FontSize").Value);
                this.Font = new Font(this.Font.Name, this.FontSize, this.Font.Style, this.Font.Unit, this.Font.GdiCharSet, this.Font.GdiVerticalFont);
                if (this.FontSize == 8)
                {
                    toolStripMenuItemFontSize8.Checked = true;
                    toolStripMenuItemFontSize10.Checked = false;
                    toolStripMenuItemFontSize12.Checked = false;
                    toolStripMenuItemFontSize14.Checked = false;
                    toolStripMenuItemFontSize16.Checked = false;
                    toolStripMenuItemFontSize18.Checked = false;
                }
                else if (this.FontSize == 10)
                {
                    toolStripMenuItemFontSize8.Checked = false;
                    toolStripMenuItemFontSize10.Checked = true;
                    toolStripMenuItemFontSize12.Checked = false;
                    toolStripMenuItemFontSize14.Checked = false;
                    toolStripMenuItemFontSize16.Checked = false;
                    toolStripMenuItemFontSize18.Checked = false;
                }
                else if (this.FontSize == 12)
                {
                    toolStripMenuItemFontSize8.Checked = false;
                    toolStripMenuItemFontSize10.Checked = false;
                    toolStripMenuItemFontSize12.Checked = true;
                    toolStripMenuItemFontSize14.Checked = false;
                    toolStripMenuItemFontSize16.Checked = false;
                    toolStripMenuItemFontSize18.Checked = false;
                }
                else if (this.FontSize == 14)
                {
                    toolStripMenuItemFontSize8.Checked = false;
                    toolStripMenuItemFontSize10.Checked = false;
                    toolStripMenuItemFontSize12.Checked = false;
                    toolStripMenuItemFontSize14.Checked = true;
                    toolStripMenuItemFontSize16.Checked = false;
                    toolStripMenuItemFontSize18.Checked = false;
                }
                else if (this.FontSize == 16)
                {
                    toolStripMenuItemFontSize8.Checked = false;
                    toolStripMenuItemFontSize10.Checked = false;
                    toolStripMenuItemFontSize12.Checked = false;
                    toolStripMenuItemFontSize14.Checked = false;
                    toolStripMenuItemFontSize16.Checked = true;
                    toolStripMenuItemFontSize18.Checked = false;
                }
                else if (this.FontSize == 18)
                {
                    toolStripMenuItemFontSize8.Checked = false;
                    toolStripMenuItemFontSize10.Checked = false;
                    toolStripMenuItemFontSize12.Checked = false;
                    toolStripMenuItemFontSize14.Checked = false;
                    toolStripMenuItemFontSize16.Checked = false;
                    toolStripMenuItemFontSize18.Checked = true;
                }
            }
            catch (Exception ex)
            {
                DisableEnableControls(false);
                StackFrame file_info = new StackFrame(true);
                Messages.error(ref file_info, ex.Message, this);
            }
        }

        void MainForm_Shown(object sender, EventArgs e)
        {
            CheckTextBoxSearchAndFillListView();

            if (Variables.CurrentUserName != "" && Variables.CurrentUserID != "")
            {
                int contactsNumbers = Variables.xDocument.Descendants("Item").Where(q => q.Attribute("UserID").Value == Variables.CurrentUserID).Count();
                this.Text = Variables.Caption + Variables.CurrentUserName + " : " + contactsNumbers.ToString() + " Contacts";
                DisableEnableControls(true);
            }
            else
                DisableEnableControls(false);
        }

        void DisableEnableControls(bool enable)
        {
            if (enable)
            {
                changeInfoToolStripMenuItem.Enabled = settingsToolStripMenuItem.Enabled = true;
                textBoxSearch.Enabled = listView1.Enabled = true;
                if (Variables.CurrentUserName == "سجاد")
                {
                    radioButton1.Enabled = true;
                    buttonEdit.Enabled = true;
                    buttonDelete.Enabled = true;
                    buttonDownloadLetter.Enabled = true;
                }
                else
                {
                    radioButton1.Enabled = false;
                    buttonEdit.Enabled = false;
                    buttonDelete.Enabled = false;
                    buttonDownloadLetter.Enabled = false;
                }
            }
            else
            {
                listView1.Items.Clear();
                changeInfoToolStripMenuItem.Enabled = settingsToolStripMenuItem.Enabled = false;
                textBoxSearch.Enabled = listView1.Enabled = false;
                buttonEdit.Enabled = false;
                buttonDownloadLetter.Enabled = false;
                buttonDelete.Enabled = false;
            }
        }

        string ConvertToPersianDate(string stringDate)
        {
            try
            {
                DateTime dateTime = DateTime.Parse(stringDate);
                PersianCalendar persianCalendar = new PersianCalendar();
                var str = persianCalendar.GetYear(dateTime).ToString() + " / " +
                                            persianCalendar.GetMonth(dateTime).ToString() + " / " +
                                            persianCalendar.GetDayOfMonth(dateTime).ToString();

                return str;
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                Messages.error(ref file_info, ex.Message, this);
                return "";
            }
        }

        string ConvertToPersianDateTime(string stringDate)
        {
            try
            {
                DateTime dateTime = DateTime.Parse(stringDate);
                PersianCalendar persianCalendar = new PersianCalendar();
                var str = persianCalendar.GetYear(dateTime).ToString() + " / " +
                                            persianCalendar.GetMonth(dateTime).ToString() + " / " +
                                            persianCalendar.GetDayOfMonth(dateTime).ToString() + "   " +
                                            persianCalendar.GetHour(dateTime).ToString() + ":" +
                                            persianCalendar.GetMinute(dateTime).ToString() + ":" +
                                            persianCalendar.GetSecond(dateTime).ToString();
                return str;
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                Messages.error(ref file_info, ex.Message, this);
                return "";
            }
        }

        #region listview

        void textBoxSearch_TextChanged(object sender, EventArgs e)
        {
            CheckTextBoxSearchAndFillListView();
        }

        private void CheckTextBoxSearchAndFillListView()
        {
            try
            {
                listView1.Items.Clear();

                var items = RecivedRepository.Instance.SearchByWord(textBoxSearch.Text.Trim());
                if (items.Count() < 1)
                    return;

                FillListView(items);
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                Messages.error(ref file_info, ex.Message, this);
            }
        }

        void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //var item = listView1.GetItemAt(e.X, e.Y);
            if (buttonEdit.Enabled)
                buttonEdit_Click(null, null);
        }

        private void SnapShotPNG(string destination)
        {
            //destination += "\\print-test.jpg";


            //Graphics canvas = listView1.CreateGraphics();
            //Bitmap bmp = new Bitmap(listView1.Width, listView1.Height, canvas);
            //listView1.DrawToBitmap(bmp, new Rectangle(0, 0, listView1.Width, listView1.Height));
            //bmp.Save(destination);



            //Creating iTextSharp Table from the DataTable data
            //iTextSharp.text.pdf.BaseFont farsiFont = iTextSharp.text.pdf.BaseFont.CreateFont(@"C:\Users\Administrator\AppData\Local\Microsoft\Windows\Fonts\B Nazanin.ttf", "UTF-8", iTextSharp.text.pdf.BaseFont.EMBEDDED);
            iTextSharp.text.pdf.BaseFont farsiFont = iTextSharp.text.pdf.BaseFont.CreateFont(@"C:\Users\Administrator\AppData\Local\Microsoft\Windows\Fonts\B Nazanin.ttf", iTextSharp.text.pdf.BaseFont.IDENTITY_H, true);
            //iTextSharp.text.pdf.BaseFont farsiFont = new iTextSharp.text.Font("B Nazanin", 16, Font.NORMAL);
            iTextSharp.text.Font paraFont = new iTextSharp.text.Font(farsiFont);
            iTextSharp.text.pdf.PdfPTable pdfTable = new iTextSharp.text.pdf.PdfPTable(listView1.Columns.Count);
            pdfTable.DefaultCell.Padding = 3;
            pdfTable.WidthPercentage = 60;
            pdfTable.HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT;
            pdfTable.DefaultCell.BorderWidth = 1;

            //Adding Header row
            foreach (ColumnHeader column in listView1.Columns)
            {
                var phrase = new iTextSharp.text.Phrase(column.Text, paraFont);
                //phrase.Font = paraFont; 

                iTextSharp.text.pdf.PdfPCell cell = new iTextSharp.text.pdf.PdfPCell(phrase);
                pdfTable.AddCell(cell);
            }

            //Adding DataRow
            foreach (ListViewItem itemRow in listView1.Items)
            {
                int i = 0;
                for (i = 0; i < itemRow.SubItems.Count - 1; i++)
                {
                    var phrase = new iTextSharp.text.Phrase(itemRow.SubItems[i].Text, paraFont);
                    //phrase.Font = paraFont;

                    iTextSharp.text.pdf.PdfPCell cell = new iTextSharp.text.pdf.PdfPCell(phrase);
                    pdfTable.AddCell(cell);
                }
            }

            //Exporting to PDF
            //string folderPath = @"D:/Temp/";
            string folderPath = destination;
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            var now = DateTime.Now;

            using (FileStream stream = new FileStream(folderPath + @$"\DataGridViewExport {now.Year}-{now.Month}-{now.Day} {now.Hour}-{now.Minute}.pdf", FileMode.Create))
            {
                iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A2, 10f, 10f, 10f, 0f);
                iTextSharp.text.pdf.PdfWriter.GetInstance(pdfDoc, stream);
                pdfDoc.Open();
                pdfDoc.Add(pdfTable);
                pdfDoc.Close();
                stream.Close();
            }
        }

        #endregion

        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == LvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (LvwColumnSorter.Order == SortOrder.Ascending)
                {
                    LvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    LvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                LvwColumnSorter.SortColumn = e.Column;
                LvwColumnSorter.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            this.listView1.Sort();

            SetColorOfTable();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            SwitchToSent();
        }

        private SentForm SwitchToSent(bool needLogin = false)
        {
            this.Hide();
            SentForm sentForm;

            if (needLogin)
                sentForm = new SentForm(needLogin);
            else
                sentForm = new SentForm();

            sentForm.Closed += (s, args) => this.Close();
            sentForm.Show();

            return sentForm;
        }

        private void buttonDownloadLetter_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count < 1) return;

            var id = Convert.ToInt32(listView1.SelectedItems[0].Name);

            var sentLetter = RecivedRepository.Instance.GetRecivedLetterById(id);

            Directory.CreateDirectory($"{_fileSharingBase}\\Softwares\\دفتر اندیکاتور\\بارگیری ها\\دریافتی");

            string fileName = Path.Combine($"{_fileSharingBase}\\Softwares\\دفتر اندیکاتور\\بارگیری ها\\دریافتی",
                listView1.SelectedItems[0].SubItems[7].Text +
                "." + sentLetter.ScanFileExtension);

            File.WriteAllBytes(fileName, sentLetter.ScanFile);

            MessageBox.Show("فایل بارگیری شد.");

            Process.Start(Environment.GetEnvironmentVariable("WINDIR") +
    @"\explorer.exe", $"{_fileSharingBase}\\Softwares\\دفتر اندیکاتور\\بارگیری ها\\دریافتی");
        }
    }
}
