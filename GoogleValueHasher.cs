using Google.Apis.Sheets.v4.Data;
using System.Globalization;
using System.Text;

namespace Sushi.Mediakiwi.Module.GoogleSheetsSync
{
    internal class GoogleValueHasher
    {
        private static readonly CultureInfo enCulture = new CultureInfo("en-US");

        public static string CreateHash(IList<CellData> cellValues)
        {
            List<string> objectValues = new List<string>();

            foreach (var item in cellValues)
            {
                objectValues.Add(GetValue(item));
            }

            return CreateMD5(string.Join(".", objectValues));
        }

        private static string GetValue(CellData input)
        {
            if (input?.UserEnteredValue?.NumberValue.HasValue == true)
            {
                if (input?.UserEnteredFormat?.NumberFormat?.Type?.Equals("DATE_TIME", StringComparison.InvariantCulture) == true)
                {
                    var dateTime = DateTime.FromOADate(input.UserEnteredValue.NumberValue.Value);
                    return dateTime.Ticks.ToString();
                }
                else
                {
                    return input.UserEnteredValue.NumberValue.Value.ToString(enCulture);
                }
            }
            else if (input?.UserEnteredValue?.BoolValue.HasValue == true)
            {
                return input.UserEnteredValue.BoolValue.Value.ToString(enCulture);
            }
            else if (string.IsNullOrWhiteSpace(input?.UserEnteredValue?.StringValue) == false)
            {
                return input.UserEnteredValue.StringValue;
            }
            else
            {
                return "";
            }
        }

        public static string CreateMD5(string input)
        {
            // Use input string to calculate MD5 hash
            using (System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create())
            {
                byte[] inputBytes = Encoding.UTF8.GetBytes(input);
                byte[] hashBytes = md5.ComputeHash(inputBytes);

                // Convert the byte array to hexadecimal string
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < hashBytes.Length; i++)
                {
                    sb.Append(hashBytes[i].ToString("X2"));
                }
                return sb.ToString();
            }
        }
    }
}
