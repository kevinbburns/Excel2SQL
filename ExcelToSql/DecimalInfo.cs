
namespace ExcelToSql
{
    public struct DecimalInfo
    {
        public int Precision { get;set; }
        public int Scale { get; set; }
        public int TrailingZeros { get; set; }

        public DecimalInfo(int precision, int scale, int trailingZeros)
            : this()
        {
            Precision     = precision;
            Scale         = scale;
            TrailingZeros = trailingZeros;
        }
    }
}
