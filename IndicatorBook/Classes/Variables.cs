using System.Xml.Linq;

namespace IndicatorBook.Classes
{
    class Variables
    {
        static public string DBFile = Application.StartupPath + "\\database";
        static public XDocument xDocument;
        static public string CurrentUserID = "";
        static public string CurrentUserName = "";
        static public string Caption = "IndicatorBook --> ";
    }
}
