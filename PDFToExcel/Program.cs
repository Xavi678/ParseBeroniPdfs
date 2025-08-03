// See https://aka.ms/new-console-template for more information
using ClosedXML.Excel;
using PDFToExcel;
using System.Globalization;
using Tabula;
using Tabula.Detectors;
using Tabula.Extractors;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;
using UglyToad.PdfPig.Fonts.Standard14Fonts;
using UglyToad.PdfPig.Geometry;
using UglyToad.PdfPig.Util;
using UglyToad.PdfPig.Writer;


NumberFormatInfo numberFormatInfo = new NumberFormatInfo();
numberFormatInfo.CurrencyDecimalSeparator = ".";
numberFormatInfo.NumberDecimalSeparator = ".";
numberFormatInfo.PercentDecimalSeparator = ".";
var fitxers = Directory.EnumerateFiles("C:\\Users\\Admin\\Documents\\Prova Beroni\\01-BSP 2025");
foreach (var fitxer in fitxers)
{
    List<string> liniesTKT = new List<string>();
    using (PdfDocument document = PdfDocument.Open(fitxer, new ParsingOptions() { ClipPaths = true }))
    {


        PageArea pageArea = ObjectExtractor.Extract(document, 3);

        // detect canditate table zones
        SimpleNurminenDetectionAlgorithm detector = new SimpleNurminenDetectionAlgorithm();
        var regions = detector.Detect(pageArea);

        IExtractionAlgorithm ea = new BasicExtractionAlgorithm();
        //var rect=UglyToad.PdfPig.Core.PdfRectangle();
        IReadOnlyList<Table> tables = ea.Extract(pageArea.GetArea(regions[0].BoundingBox)); // take first candidate area
        var table = tables[0];
        var rows = table.Rows;
        var builder = new PdfDocumentBuilder { };
        PdfDocumentBuilder.AddedFont font = builder.AddStandard14Font(Standard14Font.Helvetica);
        var paginaProva = 2;
        var pageBuilder = builder.AddPage(document, paginaProva);
        pageBuilder.SetStrokeColor(0, 255, 0);
        var outputPath = "C:\\Users\\Admin\\Documents\\Prova Beroni\\Resultat PDF Parsejat.pdf";
        List<LiniaLiquidacio> llista = new List<LiniaLiquidacio>();
        foreach (Page page in document.GetPages().Skip(1).Take(paginaProva))
        {
            var pageSegmenter = DocstrumBoundingBoxes.Instance;
            var words = page.GetWords();
            var tkttwords = words.Where(x => x.Text == "TKTT").OrderBy(x => x.BoundingBox.Top);
            var textBlocks = pageSegmenter.GetBlocks(words);

            foreach (var item in tkttwords)
            {
                var cia = ObtenirCamp(23, 32, pageBuilder, page, item);
                var num_doc = ObtenirCamp(65, 95, pageBuilder, page, item);
                var emision = ObtenirCamp(113, 136, pageBuilder, page, item);
                var cpui = ObtenirCamp(149, 163, pageBuilder, page, item);
                var stat = ObtenirCamp(199, 206, pageBuilder, page, item);
                var fop = ObtenirCamp(220, 230, pageBuilder, page, item);
                var import_transacc = ObtenirCamp(270, 289, pageBuilder, page, item);
                var tarifa = ObtenirCamp(325, 345, pageBuilder, page, item);
                var tasas = ObtenirCamp(386, 396, pageBuilder, page, item);
                var g_c = ObtenirCamp(434, 447, pageBuilder, page, item);
                //var pen = ObtenirCamp(434, 447, pageBuilder, page, item);
                var cobl = ObtenirCamp(540, 559, pageBuilder, page, item);
                var std_perc = ObtenirCamp(569, 582, pageBuilder, page, item);
                var std_import = ObtenirCamp(618, 628, pageBuilder, page, item);
                var supp_perc = ObtenirCamp(643, 654, pageBuilder, page, item);
                var supp_import = ObtenirCamp(697, 708, pageBuilder, page, item);
                var neto_pagar = ObtenirCamp(799, 816, pageBuilder, page, item);
                //var textInRegion = string.Join(" ", wordsInRegion.Select(x => x.Text).ToList());
                LiniaLiquidacio liniaLiquidacio = new LiniaLiquidacio()
                {
                    CIA = cia,
                    NUM_DOC = num_doc,
                    FECHA_EMISION = emision,
                    CPUI = cpui,
                    STAT = stat,
                    FOP = fop,
                    IMPORT_TRANSACC = ToDoubleNullable(import_transacc),
                    TARIFA = ToDoubleNullable(tarifa?.Replace("*", "")),
                    TASAS = ToDoubleNullable(tasas),
                    G_C = ToDoubleNullable(g_c),
                    COBL = ToDoubleNullable(cobl),
                    STD_IMPORTE = ToDoubleNullable(std_import),
                    STD_PERC = ToDoubleNullable(std_perc),
                    SUPP_PERC = ToDoubleNullable(supp_perc),
                    SUPP_IMPORTE = ToDoubleNullable(supp_import),
                    NETO_PAG = ToDoubleNullable(neto_pagar)
                };
                llista.Add(liniaLiquidacio);
            }

            byte[] fileBytes = builder.Build();
            File.WriteAllBytes(outputPath, fileBytes);
            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet(1);
                ws.Cell(1, 1).InsertData(llista);
                wb.SaveAs("C:\\Users\\Admin\\Documents\\Prova Beroni\\Resultat XL Parsejat.xlsx");
            }
            //IEnumerable<Word> words = page.GetWords(NearestNeighbourWordExtractor.Instance);
        }
    }

}


double? ToDoubleNullable(string? text)
{
    if (text == null)
    {
        return default;
    }

    double res;
    var isParsed = double.TryParse(text, NumberStyles.Float, numberFormatInfo, out res);
    return isParsed ? res : default(double?);
}
static string? ObtenirCamp(int bottomLeftX, int topRightX, PdfPageBuilder pageBuilder, Page page, Word item)
{
    var fy = Math.Ceiling(item.BoundingBox.TopLeft.Y);
    var fx = Math.Floor(item.BoundingBox.BottomRight.Y);
    var bottomLeft = new PdfPoint(bottomLeftX, fx);
    var topRight = new PdfPoint(topRightX, fy);
    var square = new PdfRectangle(bottomLeft, topRight);
    pageBuilder.DrawRectangle(square.BottomLeft, square.Width, square.Height);
    pageBuilder.DrawRectangle(item.BoundingBox.BottomLeft, item.BoundingBox.Width, item.BoundingBox.Height);
    var letters = page.Letters.Where(x => square.IntersectsWith(x.GlyphRectangle)).ToList();

    var wordsInRegion = DefaultWordExtractor.Instance.GetWords(letters);
    return wordsInRegion.FirstOrDefault()?.Text;
}