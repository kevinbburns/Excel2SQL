using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDataReader;

namespace ExcelToSql
{
    public static class Program
    {
        private const string SafeNumber = "255";

        private static readonly Dictionary<Type, string> TypesRef = new Dictionary<Type, string>
        {
            { typeof(string), "nvarchar"},
            {typeof(Guid), "uniqueidentifier"},
            {typeof(long), "bigint"},
            {typeof(byte[]), "binary"},
            {typeof(bool), "bit"},
            {typeof(DateTime), "datetime"},
            {typeof(decimal), "decimal"},
            {typeof(double), "float"},
            {typeof(int), "int"},
            {typeof(float), "real"},
            {typeof(short), "smallint"},
            {typeof(byte), "tinyint"},
            {typeof(object), "nvarchar(512)"},
            {typeof(DateTimeOffset), "datetimeoffset"},
        };

        private static void Main(string[] args)
        {

            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                if (args.Count() < 4 || (!args.Contains("/I") || !args.Contains("/O")))
                {
                    PrintHelp();
                    return;
                }
               
                var inputFile          = GetArgumentParameter("/I", args);
                var outputFile         = GetArgumentParameter("/O", args);
                var inFile             = new FileInfo(inputFile);

                if (string.IsNullOrEmpty(inputFile)  ||
                    string.IsNullOrEmpty(outputFile) || 
                    !inFile.Exists)
                {
                    PrintHelp();
                    return;
                }

                var firstRowHeaders    = GetArgumentBoolean("/H", args);
                var schema             = GetArgumentParameter("/D", args);
                var identityInsert     = GetArgumentBoolean("/K", args);

                if (string.IsNullOrEmpty(schema))
                    schema             = "dbo";

                var sqlString          = CreateSql(inFile, firstRowHeaders,identityInsert,schema);

                try
                {
                    using (var fStream = File.OpenWrite(outputFile))
                    using (var writer  = new StreamWriter(fStream))
                        writer.Write(sqlString);
                    Console.WriteLine("File successfully created.");
                    Console.ReadLine();
                }
                catch (IOException i)
                {
                    Console.WriteLine(i.Message);
                    Console.ReadLine();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    Console.ReadLine();
                }
            }
            catch (IOException)
            {
                Console.WriteLine("The file is open in an another process, please check for open Excel windows.");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private static void PrintHelp()
        {
            Console.WriteLine("\n Example: Excel2Sql /I input-file.xlsx /O input-file.sql");
            Console.WriteLine("\n Optional Parameters:\n");
            Console.WriteLine(" /H Include this if the first row contains headers.\n");
            Console.WriteLine(" /K Include this if you want to include an integer identity column.\n");
            //Console.WriteLine(" /P Generates a C# DTO/POCO, must supply valid path/name for *.cs file.");
            Console.WriteLine(" /D Schema (dbo is used by default)");
            Console.ReadLine();
        }

        private static string GetArgumentParameter(string optionName, IReadOnlyList<string> args)
        {
            var cnt      = 0;
            foreach (var arg in args)
            {
                if (string.Equals(arg, optionName, StringComparison.CurrentCultureIgnoreCase))
                {
                    return args[cnt + 1];
                }
                cnt += 1;
            }

            return string.Empty;
        }

        private static bool GetArgumentBoolean(string optionName, IEnumerable<string> args)
        {
            var rArg     = string.Empty;
            foreach (var arg in args)
            {
                if (string.Equals(arg, optionName, StringComparison.CurrentCultureIgnoreCase))
                {
                    rArg = arg;
                }
            }

            return rArg != string.Empty;
        }

        private static string CreateSql(FileSystemInfo inputFileInfo, bool firstRowHeaders
                                       ,bool identityInsert, string schema = "dbo")
        {
            var sb                 = new StringBuilder();
            using (var stream      = File.OpenRead(inputFileInfo.FullName))
            using (var reader      = ExcelReaderFactory.CreateReader(stream))
            using (var result      = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                UseColumnDataType  = true,
                ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration
                {
                    UseHeaderRow   = firstRowHeaders
                },
            }))
            {
                foreach (DataTable dataTable in result.Tables)
                {
                    sb.Append($"CREATE TABLE [{schema}].[{dataTable.TableName}] (");
                    if (identityInsert)
                        sb.Append("[Id] [int] IDENTITY(1,1) NOT NULL,");

                    var cnt = 0;
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        if (TypesRef.ContainsKey(column.DataType))
                        {
                            sb.Append(cnt + 1 < dataTable.Columns.Count
                                ? $"[{column.ColumnName}] {GetSqlType(column.DataType, dataTable, column)},"
                                : $"[{column.ColumnName}] {GetSqlType(column.DataType, dataTable, column)}");
                        }
                        else
                        {
                            throw new NullReferenceException();
                        }

                        cnt += 1;
                    }

                    if (!identityInsert)
                    {
                        sb.Append(")");
                    }
                    else
                    {
                        sb.Append($"CONSTRAINT [PK_{dataTable.TableName}] PRIMARY KEY CLUSTERED ( [Id] ASC ) WITH " +
                                  $"(PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, " +
                                  $"ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] ) ON [PRIMARY]");
                    }
                }
            }

            return sb.ToString();
        }

        private static string GetSqlType(Type cType,DataTable table,DataColumn column)
        {
            if (cType == typeof(string))
                return HandleStringType(table, column);

            if (cType == typeof(decimal))
                return GetDecimalMantissa(table, column);

            return ParseAndCanBeNull(table, column);
        }

        private static string ParseAndCanBeNull( DataTable table, DataColumn column)
        {
            var isNull = false;
            foreach (DataRow row in table.Rows)
            {
                if (row[column] != DBNull.Value && !string.IsNullOrEmpty(row[column].ToString()) &&
                    !string.IsNullOrWhiteSpace(row[column].ToString())) continue;
                isNull = true;
                break;
            }

            return isNull ? $"{TypesRef[column.DataType]} NULL" : $"{TypesRef[column.DataType]}";
        }

        #region String
        private static string HandleStringType(DataTable table, DataColumn column)
        {
            var max = GetMaxLengthAndIsNull(table, column.Ordinal, out var isNull);
            var strType = $"{TypesRef[column.DataType]}({max})";
            strType += (isNull ? " NULL" : "");
            return strType;
        }

        private static string GetMaxLengthAndIsNull(DataTable table, int column, out bool canBeNull)
        {
            var max = 0;
            if (table.Rows.Count <= 0)
            {
                canBeNull = true;
                return SafeNumber;
            }

            var isNull = false;
            foreach (DataRow row in table.Rows)
            {
                if (row[column] == DBNull.Value ||
                    string.IsNullOrEmpty(row[column].ToString()) ||
                    string.IsNullOrWhiteSpace(row[column].ToString()))
                {
                    isNull = true;
                    continue;
                }
                var lngth = row[column].ToString().Length;
                if (lngth > max)
                    max = lngth;
            }

            canBeNull = isNull;
            return GetSqlMax(max);
        }

        private static string GetSqlMax(int intMax)
        {
            if (intMax <= 255)
                return "255";
            if (intMax > 255 && intMax <= 512)
                return "512";
            return "MAX";
        }
        #endregion

        #region Decimal
        private static string GetDecimalMantissa(DataTable table, DataColumn column)
        {
            var max     = MaxDecimalInfo(table, column.Ordinal, out var isNull);
            var strType = $"{TypesRef[column.DataType]}{max}";
            strType += (isNull ? " NULL" : "");
            return strType;
        }

        private static string MaxDecimalInfo(DataTable table, int column, out bool canBeNull)
        {
            var info                   = new DecimalInfo();

            if (table.Rows.Count <= 0)
            {
                canBeNull              = true;
                return SafeNumber;
            }

            var isNull                 = false;
            foreach (DataRow row in table.Rows)
            {
                if (row[column] == DBNull.Value ||
                    string.IsNullOrEmpty(row[column].ToString()) ||
                    string.IsNullOrWhiteSpace(row[column].ToString()))
                {
                    isNull             = true;
                    continue;
                }
                var fastInfo           = GetFastInfo((decimal)row[column]);

                if (fastInfo.Precision > info.Precision)
                    info.Precision     = fastInfo.Precision;
                if (fastInfo.Scale > info.Scale)
                    info.Scale         = fastInfo.Scale;
                if (fastInfo.TrailingZeros > info.TrailingZeros)
                    info.TrailingZeros = fastInfo.TrailingZeros;
            }

            canBeNull                  = isNull;
            return $"({info.Precision},{info.Scale})";
        }


        //https://stackoverflow.com/questions/763942/calculate-system-decimal-precision-and-scale
        private static DecimalInfo GetFastInfo(decimal dec)
        {
            var s                       = dec.ToString(CultureInfo.InvariantCulture);

            var precision               = 0;
            var scale                   = 0;
            var trailingZeros           = 0;
            var inFraction              = false;
            var nonZeroSeen             = false;

            foreach (var c in s)
            {
                if (inFraction)
                {
                    if (c == '0')
                        trailingZeros++;
                    else
                    {
                        nonZeroSeen     = true;
                        trailingZeros   = 0;
                    }

                    precision++;
                    scale++;
                }
                else
                {
                    if (c == '.')
                    {
                        inFraction      = true;
                    }
                    else if (c != '-')
                    {
                        if (c != '0' || nonZeroSeen)
                        {
                            nonZeroSeen = true;
                            precision++;
                        }
                    }
                }
            }

            // Handles cases where all digits are zeros.
            if (!nonZeroSeen)
                precision += 1;

            return new DecimalInfo(precision, scale, trailingZeros);
        }
        #endregion
    }
}
