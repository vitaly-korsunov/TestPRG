using System.Data;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using Microsoft.Data.SqlClient;

if (args.Length == 0 || HasOption(args, "--help", "-h", "/?"))
{
    PrintUsage();
    return 0;
}

try
{
    var options = CommandLineOptions.Parse(args);
    var table = await StoredProcedureRunner.ExecuteAsync(options);

    ExcelWorkbookWriter.Write(table, options.OutputPath, options.SheetName);

    Console.WriteLine($"Exported {table.Rows.Count} rows and {table.Columns.Count} columns to '{Path.GetFullPath(options.OutputPath)}'.");
    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine("Export failed:");
    Console.Error.WriteLine(ex.Message);
    return 1;
}

static bool HasOption(string[] args, params string[] names) =>
    args.Any(arg => names.Contains(arg, StringComparer.OrdinalIgnoreCase));

static void PrintUsage()
{
    Console.WriteLine("""
    Export a SQL Server stored procedure result to an Excel .xlsx file.

    Usage:
      dotnet run -- --connection "<connection-string>" --sp dbo.ExportCustomers --out customers.xlsx

    Options:
      --connection, -c   SQL Server connection string. Required.
      --sp, -s           Stored procedure name. Required.
      --out, -o          Output .xlsx file path. Required.
      --sheet            Worksheet name. Default: Data.
      --timeout          Command timeout in seconds. Default: 120.
      --param, -p        Stored procedure parameter as Name=Value. Can be repeated.

    Example with parameters:
      dotnet run -- -c "Server=.;Database=MyDb;Trusted_Connection=True;TrustServerCertificate=True" -s dbo.ExportOrders -o orders.xlsx -p StartDate=2026-01-01 -p EndDate=2026-06-29
    """);
}

internal sealed record CommandLineOptions(
    string ConnectionString,
    string StoredProcedure,
    string OutputPath,
    string SheetName,
    int TimeoutSeconds,
    IReadOnlyList<SqlParameter> Parameters)
{
    public static CommandLineOptions Parse(string[] args)
    {
        string? connectionString = null;
        string? storedProcedure = null;
        string? outputPath = null;
        string sheetName = "Data";
        int timeoutSeconds = 120;
        var parameters = new List<SqlParameter>();

        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            switch (arg.ToLowerInvariant())
            {
                case "--connection":
                case "-c":
                    connectionString = ReadValue(args, ref i, arg);
                    break;
                case "--sp":
                case "-s":
                    storedProcedure = ReadValue(args, ref i, arg);
                    break;
                case "--out":
                case "-o":
                    outputPath = ReadValue(args, ref i, arg);
                    break;
                case "--sheet":
                    sheetName = ReadValue(args, ref i, arg);
                    break;
                case "--timeout":
                    if (!int.TryParse(ReadValue(args, ref i, arg), NumberStyles.None, CultureInfo.InvariantCulture, out timeoutSeconds) ||
                        timeoutSeconds <= 0)
                    {
                        throw new ArgumentException("--timeout must be a positive whole number.");
                    }
                    break;
                case "--param":
                case "-p":
                    parameters.Add(ParseParameter(ReadValue(args, ref i, arg)));
                    break;
                default:
                    throw new ArgumentException($"Unknown argument '{arg}'. Use --help to see valid options.");
            }
        }

        if (string.IsNullOrWhiteSpace(connectionString))
        {
            throw new ArgumentException("Missing required option --connection.");
        }

        if (string.IsNullOrWhiteSpace(storedProcedure))
        {
            throw new ArgumentException("Missing required option --sp.");
        }

        if (string.IsNullOrWhiteSpace(outputPath))
        {
            throw new ArgumentException("Missing required option --out.");
        }

        if (!outputPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            outputPath += ".xlsx";
        }

        return new CommandLineOptions(
            connectionString,
            storedProcedure,
            outputPath,
            SanitizeSheetName(sheetName),
            timeoutSeconds,
            parameters);
    }

    private static string ReadValue(string[] args, ref int index, string optionName)
    {
        if (index + 1 >= args.Length || args[index + 1].StartsWith('-'))
        {
            throw new ArgumentException($"{optionName} requires a value.");
        }

        index++;
        return args[index];
    }

    private static SqlParameter ParseParameter(string raw)
    {
        var equalsIndex = raw.IndexOf('=');
        if (equalsIndex <= 0)
        {
            throw new ArgumentException($"Parameter '{raw}' must be in Name=Value format.");
        }

        var name = raw[..equalsIndex].Trim();
        var valueText = raw[(equalsIndex + 1)..].Trim();
        if (!name.StartsWith('@'))
        {
            name = "@" + name;
        }

        return new SqlParameter(name, InferParameterValue(valueText));
    }

    private static object InferParameterValue(string value)
    {
        if (value.Equals("null", StringComparison.OrdinalIgnoreCase))
        {
            return DBNull.Value;
        }

        if (bool.TryParse(value, out var boolValue))
        {
            return boolValue;
        }

        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var intValue))
        {
            return intValue;
        }

        if (long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var longValue))
        {
            return longValue;
        }

        if (decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out var decimalValue))
        {
            return decimalValue;
        }

        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var dateValue))
        {
            return dateValue;
        }

        return value;
    }

    private static string SanitizeSheetName(string sheetName)
    {
        var invalidChars = new HashSet<char>(['[', ']', ':', '*', '?', '/', '\\']);
        var cleaned = new string(sheetName.Where(ch => !invalidChars.Contains(ch)).ToArray()).Trim();

        if (string.IsNullOrWhiteSpace(cleaned))
        {
            cleaned = "Data";
        }

        return cleaned.Length > 31 ? cleaned[..31] : cleaned;
    }
}

internal static class StoredProcedureRunner
{
    public static async Task<DataTable> ExecuteAsync(CommandLineOptions options)
    {
        await using var connection = new SqlConnection(options.ConnectionString);
        await using var command = new SqlCommand(options.StoredProcedure, connection)
        {
            CommandType = CommandType.StoredProcedure,
            CommandTimeout = options.TimeoutSeconds
        };

        foreach (var parameter in options.Parameters)
        {
            command.Parameters.Add(parameter);
        }

        await connection.OpenAsync();
        await using var reader = await command.ExecuteReaderAsync();

        var table = new DataTable(options.SheetName);
        table.Load(reader);
        return table;
    }
}

internal static class ExcelWorkbookWriter
{
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

    public static void Write(DataTable table, string outputPath, string sheetName)
    {
        var directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        if (File.Exists(outputPath))
        {
            File.Delete(outputPath);
        }

        using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Create);

        AddTextEntry(archive, "[Content_Types].xml", BuildContentTypesXml());
        AddTextEntry(archive, "_rels/.rels", BuildRootRelationshipsXml());
        AddTextEntry(archive, "xl/workbook.xml", BuildWorkbookXml(sheetName));
        AddTextEntry(archive, "xl/_rels/workbook.xml.rels", BuildWorkbookRelationshipsXml());
        AddTextEntry(archive, "xl/styles.xml", BuildStylesXml());
        AddTextEntry(archive, "xl/worksheets/sheet1.xml", BuildWorksheetXml(table));
    }

    private static void AddTextEntry(ZipArchive archive, string entryName, string content)
    {
        var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    private static string BuildContentTypesXml() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
        </Types>
        """;

    private static string BuildRootRelationshipsXml() =>
        $"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="{PackageRelationshipsNamespace}">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """;

    private static string BuildWorkbookXml(string sheetName) =>
        $"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="{SpreadsheetNamespace}" xmlns:r="{RelationshipsNamespace}">
          <sheets>
            <sheet name="{XmlEscape(sheetName)}" sheetId="1" r:id="rId1"/>
          </sheets>
        </workbook>
        """;

    private static string BuildWorkbookRelationshipsXml() =>
        $"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="{PackageRelationshipsNamespace}">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        </Relationships>
        """;

    private static string BuildStylesXml() =>
        $"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="{SpreadsheetNamespace}">
          <fonts count="2">
            <font><sz val="11"/><name val="Calibri"/></font>
            <font><b/><sz val="11"/><name val="Calibri"/></font>
          </fonts>
          <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
          <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
          <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
          <cellXfs count="2">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
            <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
          </cellXfs>
          <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
        </styleSheet>
        """;

    private static string BuildWorksheetXml(DataTable table)
    {
        using var stringWriter = new StringWriter(CultureInfo.InvariantCulture);
        using var writer = XmlWriter.Create(stringWriter, new XmlWriterSettings
        {
            OmitXmlDeclaration = false,
            Encoding = Encoding.UTF8,
            Indent = true
        });

        writer.WriteStartDocument(true);
        writer.WriteStartElement("worksheet", SpreadsheetNamespace);
        writer.WriteStartElement("cols");
        foreach (var width in CalculateColumnWidths(table))
        {
            var columnIndex = width.Index + 1;
            writer.WriteStartElement("col");
            writer.WriteAttributeString("min", columnIndex.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("max", columnIndex.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("width", width.Width.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("customWidth", "1");
            writer.WriteEndElement();
        }
        writer.WriteEndElement();

        writer.WriteStartElement("sheetData");

        WriteHeaderRow(writer, table);
        WriteDataRows(writer, table);

        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndDocument();
        writer.Flush();

        return stringWriter.ToString();
    }

    private static void WriteHeaderRow(XmlWriter writer, DataTable table)
    {
        writer.WriteStartElement("row");
        writer.WriteAttributeString("r", "1");

        for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
        {
            WriteInlineStringCell(writer, 1, columnIndex, table.Columns[columnIndex].ColumnName, styleIndex: 1);
        }

        writer.WriteEndElement();
    }

    private static void WriteDataRows(XmlWriter writer, DataTable table)
    {
        for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            var excelRowIndex = rowIndex + 2;
            writer.WriteStartElement("row");
            writer.WriteAttributeString("r", excelRowIndex.ToString(CultureInfo.InvariantCulture));

            for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
            {
                var value = table.Rows[rowIndex][columnIndex];
                WriteCell(writer, excelRowIndex, columnIndex, value);
            }

            writer.WriteEndElement();
        }
    }

    private static void WriteCell(XmlWriter writer, int rowIndex, int columnIndex, object value)
    {
        if (value is null || value == DBNull.Value)
        {
            writer.WriteStartElement("c");
            writer.WriteAttributeString("r", GetCellReference(rowIndex, columnIndex));
            writer.WriteEndElement();
            return;
        }

        switch (value)
        {
            case byte or short or int or long or float or double or decimal:
                WriteNumberCell(writer, rowIndex, columnIndex, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                break;
            case bool boolValue:
                WriteBooleanCell(writer, rowIndex, columnIndex, boolValue);
                break;
            case DateTime dateTime:
                WriteInlineStringCell(writer, rowIndex, columnIndex, dateTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
                break;
            default:
                WriteInlineStringCell(writer, rowIndex, columnIndex, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                break;
        }
    }

    private static void WriteInlineStringCell(XmlWriter writer, int rowIndex, int columnIndex, string value, int? styleIndex = null)
    {
        writer.WriteStartElement("c");
        writer.WriteAttributeString("r", GetCellReference(rowIndex, columnIndex));
        writer.WriteAttributeString("t", "inlineStr");
        if (styleIndex is not null)
        {
            writer.WriteAttributeString("s", styleIndex.Value.ToString(CultureInfo.InvariantCulture));
        }

        writer.WriteStartElement("is");
        writer.WriteElementString("t", value);
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteNumberCell(XmlWriter writer, int rowIndex, int columnIndex, string value)
    {
        writer.WriteStartElement("c");
        writer.WriteAttributeString("r", GetCellReference(rowIndex, columnIndex));
        writer.WriteStartElement("v");
        writer.WriteString(value);
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteBooleanCell(XmlWriter writer, int rowIndex, int columnIndex, bool value)
    {
        writer.WriteStartElement("c");
        writer.WriteAttributeString("r", GetCellReference(rowIndex, columnIndex));
        writer.WriteAttributeString("t", "b");
        writer.WriteStartElement("v");
        writer.WriteString(value ? "1" : "0");
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static IEnumerable<(int Index, int Width)> CalculateColumnWidths(DataTable table)
    {
        for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
        {
            var maxLength = table.Columns[columnIndex].ColumnName.Length;
            foreach (DataRow row in table.Rows)
            {
                var value = row[columnIndex] == DBNull.Value ? string.Empty : Convert.ToString(row[columnIndex], CultureInfo.InvariantCulture) ?? string.Empty;
                maxLength = Math.Max(maxLength, value.Length);
            }

            yield return (columnIndex, Math.Clamp(maxLength + 2, 10, 60));
        }
    }

    private static string GetCellReference(int rowIndex, int zeroBasedColumnIndex) =>
        GetColumnName(zeroBasedColumnIndex) + rowIndex.ToString(CultureInfo.InvariantCulture);

    private static string GetColumnName(int zeroBasedColumnIndex)
    {
        var dividend = zeroBasedColumnIndex + 1;
        var columnName = new StringBuilder();

        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            columnName.Insert(0, (char)('A' + modulo));
            dividend = (dividend - modulo) / 26;
        }

        return columnName.ToString();
    }

    private static string XmlEscape(string value) => SecurityElementEscape(value);

    private static string SecurityElementEscape(string value)
    {
        return value
            .Replace("&", "&amp;", StringComparison.Ordinal)
            .Replace("<", "&lt;", StringComparison.Ordinal)
            .Replace(">", "&gt;", StringComparison.Ordinal)
            .Replace("\"", "&quot;", StringComparison.Ordinal)
            .Replace("'", "&apos;", StringComparison.Ordinal);
    }
}
