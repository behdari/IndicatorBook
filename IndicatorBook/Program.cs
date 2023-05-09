using IndicatorBook.Sent;

namespace IndicatorBook
{
    static class Program
    {
        /// <summary>
        /// The code is take from here:
        /// https://www.codeproject.com/Articles/38246/Phone-Book-in-C
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new SentForm(true));
        }
    }
}
