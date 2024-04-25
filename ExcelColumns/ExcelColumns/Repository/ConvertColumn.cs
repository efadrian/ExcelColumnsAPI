using Excel.Interface;

namespace Excel.Repository
{
    public class ConvertColumn : IConvertColumn
    {
        //
        public int ConvertColumnLetterToColumnNumber(string columnTitle)
        {

            int resultNumber = 0;
            columnTitle = columnTitle.ToUpper();
            foreach (char c in columnTitle)
            {
                resultNumber *= 26; // number * no of letters

                int value = c - 'A' + 1;
                resultNumber += value;
            }

            return resultNumber;
        }

        //
        public string GetColumnNumberToTitleColumns(int number)
        {
            string resultValue = "";
            while (number > 0)
            {
                int colNumber = (number - 1) % 26; // number divided to no of letters
                resultValue = (char)('A' + colNumber) + resultValue;
                number = (number - 1) / 26;
            }

            return resultValue;
        }

    }
}
