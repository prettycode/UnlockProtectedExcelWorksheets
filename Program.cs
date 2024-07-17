using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;

const string WORKSHEET_REGEX_PATTERN = @"xl/worksheets/.*\.xml$";
const bool DEBUG = true;

string[] files = [
    @"..\..\..\Unprotected.xlsx",
    @"..\..\..\Protected.xlsx"
];

foreach (string xlsx in files)
{
    using var zipToOpen = new FileStream(xlsx, FileMode.Open);
    using var archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update);

    var worksheets = archive.Entries.Where(entry => Regex.IsMatch(entry.FullName, WORKSHEET_REGEX_PATTERN)).ToList();
    var worksheetsCount = worksheets.Count;

    Console.WriteLine($"File: {Path.GetFullPath(xlsx)}");
    Console.WriteLine($"Found {worksheetsCount} worksheets.");

    worksheets.ForEach(sheet => RemoveSheetProtection(sheet, DEBUG));

    Console.WriteLine();
}

static void RemoveSheetProtection(ZipArchiveEntry entry, bool reportProtectionOnly)
{
    using Stream entryStream = entry.Open();
    var doc = XDocument.Load(entryStream);
    var sheetProtectionElements = doc.Descendants().Where(e => e.Name.LocalName == "sheetProtection");
    var elementCount = sheetProtectionElements.Count();

    if (elementCount < 1)
    {
        Console.WriteLine($"{entry.FullName}: No protection found.");
        return;
    }

    if (!reportProtectionOnly)
    {
        sheetProtectionElements.Remove();
        entryStream.SetLength(0);
        entryStream.Position = 0;
        doc.Save(entryStream);

        Console.WriteLine($"{entry.FullName}: Protection removed.");
    } else
    {
        Console.WriteLine($"{entry.FullName}: Protected found.");
    }
}