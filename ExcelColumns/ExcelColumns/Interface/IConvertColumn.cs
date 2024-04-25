namespace Excel.Interface
{
    public interface IConvertColumn
    {
        public  int ConvertColumnLetterToColumnNumber(string columnName);

        public string GetColumnNumberToTitleColumns(int number);
    }
}
