using System.ComponentModel;
using System.Reflection;

namespace IndicatorBook.Resources
{
    public static class IndicatorBookUtility
    {
        public static T GetEnumValueFromDescription<T>(string description)
        {
            MemberInfo[] fis = typeof(T).GetFields();

            foreach (var fi in fis)
            {
                DescriptionAttribute[] attributes = (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);

                if (attributes != null && attributes.Length > 0 && attributes[0].Description == description)
                    return (T)Enum.Parse(typeof(T), fi.Name);
            }

            throw new Exception("Not found");
        }

        public static string GetDescription<T>(this T enumValue)
            where T : struct, IConvertible
        {
            if (!typeof(T).IsEnum)
                return null;

            var description = enumValue.ToString();
            var fieldInfo = enumValue.GetType().GetField(enumValue.ToString());

            if (fieldInfo != null)
            {
                var attrs = fieldInfo.GetCustomAttributes(typeof(DescriptionAttribute), true);
                if (attrs != null && attrs.Length > 0)
                {
                    description = ((DescriptionAttribute)attrs[0]).Description;
                }
            }

            return description;
        }
    }
}
