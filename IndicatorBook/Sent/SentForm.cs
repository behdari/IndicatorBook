using ClosedXML.Excel;
using IndicatorBook.Classes;
using IndicatorBook.Recived;
using IndicatorBook.Resources;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Xml.Linq;
using SaveOptions = System.Xml.Linq.SaveOptions;

namespace IndicatorBook.Sent
{
    public partial class SentForm : Form
    {
        float FontSize = 10.0f;
        private bool isInit = true;
        private bool _isLogin;
        public string _fileSharingBase;

        public ListViewColumnSorter LvwColumnSorter { get; private set; }

        public SentForm(bool isLogin = false)
        {
            _isLogin = isLogin;
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
                int previousNumber = 0;
                string previousFile = "";
                string previousPamphleteer = "";

                if (listView1.Items.Count > 0)
                {
                    previousNumber = Convert.ToInt32(listView1.Items[0].SubItems[5].Text);
                    previousFile = listView1.Items[0].SubItems[7].Text;
                    previousPamphleteer = listView1.Items[0].SubItems[6].Text;
                }

                SentItemForm newForm = new SentItemForm(true, false, _fileSharingBase);
                newForm.Font = new Font(this.Font.Name, this.FontSize, this.Font.Style, this.Font.Unit, this.Font.GdiCharSet, this.Font.GdiVerticalFont);
                newForm.Text = "افزودن آیتم جدید";
                newForm.persianDatePickerDate.GeorgianDate = DateTime.Today;
                newForm.textBoxFile.Text = previousFile;
                newForm.textBoxPamphleteer.Text = previousPamphleteer;
                newForm.textBoxNumber.Text = (previousNumber + 1).ToString();
                newForm.textBoxNumber.Text = (previousNumber + 1).ToString();
                newForm.textBoxNextNumber.Text = (previousNumber + 2).ToString();

                //newForm.lableRegDate.Text = christianToolStripMenuItem.Checked ? DateTime.Now.ToString() : ConvertToPersianDate(DateTime.Now.ToString());
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

                SentItemForm editForm = new SentItemForm(false, true, _fileSharingBase);

                editForm.Font = new Font(this.Font.Name, this.FontSize, this.Font.Style, this.Font.Unit, this.Font.GdiCharSet, this.Font.GdiVerticalFont);
                editForm.Text = "ویرایش آیتم";

                editForm.textBoxPamphleteer.Text = listView1.SelectedItems[0].SubItems[7].Text;
                editForm.textBoxFile.Text = listView1.SelectedItems[0].SubItems[6].Text;
                editForm.textBoxNumber.Text = listView1.SelectedItems[0].SubItems[5].Text;
                editForm.textBoxTitle.Text = listView1.SelectedItems[0].SubItems[4].Text;
                editForm.textBoxDescription.Text = listView1.SelectedItems[0].SubItems[3].Text;
                editForm.textBoxNextNumber.Text = listView1.SelectedItems[0].SubItems[2].Text;

                var hasAttachment = listView1.SelectedItems[0].SubItems[1].Text == "دارد";
                if (hasAttachment)
                {
                    editForm.radioButtonHasAttachment.Checked = true;
                }
                else
                {
                    editForm.radioButtonHasNotAttachment.Checked = true;
                }

                var splitedDate = listView1.SelectedItems[0].SubItems[0].Text.Split('/');
                editForm.persianDatePickerDate.PersianDate = new PersianDateInfo(Convert.ToInt32(splitedDate[0].Trim()), // year
                                                                                 Convert.ToInt32(splitedDate[1].Trim()), // month
                                                                                 Convert.ToInt32(splitedDate[2].Trim())); // day

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

                SentRepository.Instance.Delete(id);

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

        public void newUserToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                UserForm newUserForm = new UserForm(true, false, false, isInit);
                newUserForm.Font = new Font(this.Font.Name, this.FontSize, this.Font.Style, this.Font.Unit, this.Font.GdiCharSet, this.Font.GdiVerticalFont);
                newUserForm.ShowDialog();
                ApplySettings();
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
            catch (Exception ex)
            {
                DisableEnableControls(false);
                StackFrame file_info = new StackFrame(true);
                Messages.error(ref file_info, ex.Message, this);
            }
        }

        void changeUserToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                UserForm userForm = new UserForm(false, true, false, isInit);
                userForm.Font = new Font(this.Font.Name, this.FontSize, this.Font.Style, this.Font.Unit, this.Font.GdiCharSet, this.Font.GdiVerticalFont);
                userForm.ShowDialog();
                ApplySettings();
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
            catch (Exception ex)
            {
                DisableEnableControls(false);
                StackFrame file_info = new StackFrame(true);
                Messages.error(ref file_info, ex.Message, this);
            }
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

        private void FillListView(List<SentLetter> items)
        {
            foreach (var item in items)
            {
                ListViewItem listViewItems;

                listViewItems = new ListViewItem(new string[]
                                { item.Date.ToString(),
                                  item.HasAttachment? "دارد" : "ندارد",
                                  item.NextNumber.ToString(),
                                  item.Description,
                                  item.Title,
                                  item.Number.ToString(),
                                  item.Pamphleteer,
                                  item.File });

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
                    listView1.Items[i].BackColor = Color.FromArgb(229, 255, 204);
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
            if (_isLogin)
            {
                try
                {
                    if (!File.Exists(Variables.DBFile))
                    {
                        newUserToolStripMenuItem_Click(null, null);
                        isInit = false;
                        return;
                    }

                    Variables.xDocument = XDocument.Parse(TripleDES.DecryptFromFile(Variables.DBFile, TripleDES.ByteKey, TripleDES.IV));

                    var users = from q in Variables.xDocument.Descendants("User")
                                select q;

                    if (users.Count() < 1)//No user exist
                    {
                        newUserToolStripMenuItem_Click(null, null);
                        isInit = false;
                        return;
                    }
                    else//More than one user exist
                    {
                        changeUserToolStripMenuItem_Click(null, null);
                    }

                    isInit = false;
                }
                catch (Exception ex)
                {
                    DisableEnableControls(false);
                    StackFrame file_info = new StackFrame(true);
                    Messages.error(ref file_info, ex.Message, this);
                    try
                    {
                        //File.Delete(Variables.DBFile);
                    }
                    catch
                    {
                        MessageBox.Show("Please delete the DataBase file", "Fatal Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
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
                    buttonDownloadLetter.Enabled = true;
                    buttonDelete.Enabled = true;
                }
                else
                {
                    radioButton1.Enabled = false;
                    buttonDownloadLetter.Enabled = false;
                    buttonDelete.Enabled = false;
                }
            }
            else
            {
                listView1.Items.Clear();
                changeInfoToolStripMenuItem.Enabled = settingsToolStripMenuItem.Enabled = false;
                textBoxSearch.Enabled = listView1.Enabled = false;
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

                var items = SentRepository.Instance.SearchByWord(textBoxSearch.Text.Trim());
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
            if (buttonDownloadLetter.Enabled)
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
            this.Hide();
            RecivedForm recivedForm = new RecivedForm();
            recivedForm.Closed += (s, args) => this.Close();
            recivedForm.Show();
        }

        private void buttonReciveRawFile_Click(object sender, EventArgs e)
        {
            var rawFileForm = new SentTemplateForm(_fileSharingBase);

            rawFileForm.ShowDialog();
        }

        private void buttonDownloadLetter_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count < 1) return;

            var id = Convert.ToInt32(listView1.SelectedItems[0].Name);

            var sentLetter = SentRepository.Instance.GetSentLetterById(id);

            Directory.CreateDirectory($"{_fileSharingBase}\\Softwares\\دفتر اندیکاتور\\بارگیری ها\\ارسالی");

            string fileName = Path.Combine($"{_fileSharingBase}\\Softwares\\دفتر اندیکاتور\\بارگیری ها\\ارسالی",
                listView1.SelectedItems[0].SubItems[6].Text.Replace("/", "-") + " " +
                listView1.SelectedItems[0].SubItems[7].Text + " " +
                listView1.SelectedItems[0].SubItems[5].Text + " " +
                "." + sentLetter.WordFileExtension);

            File.WriteAllBytes(fileName, sentLetter.WordFile);

            MessageBox.Show("فایل بارگیری شد.");

            Process.Start(Environment.GetEnvironmentVariable("WINDIR") +
    @"\explorer.exe", $"{_fileSharingBase}\\Softwares\\دفتر اندیکاتور\\بارگیری ها\\ارسالی");
        }
    }
}
