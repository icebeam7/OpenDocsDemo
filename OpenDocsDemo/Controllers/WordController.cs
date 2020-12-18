using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using OpenDocsDemo.Helpers;

using Microsoft.AspNetCore.Mvc;

namespace OpenDocsDemo.Controllers
{
    public class WordController : Controller
    {
        string docxMIMEType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult CreateDocument()
        {
            using (var stream = new MemoryStream())
            {
                using (var wordDocument = WordprocessingDocument.Create(stream,
                    WordprocessingDocumentType.Document, true))
                {
                    wordDocument.AddMainDocumentPart();

                    var document = new Document();
                    var body = new Body();

                    var paragraph = new Paragraph();

                    var paragraphProperties = new ParagraphProperties();
                    var paragraphStyleId = new ParagraphStyleId() { Val = "Normal" };
                    var centerJustification = new Justification() { Val = JustificationValues.Center };

                    paragraphProperties.Append(paragraphStyleId);
                    paragraphProperties.Append(centerJustification);

                    var run = new Run();
                    var text = new Text("Hello world from Open XML SDK!");

                    run.Append(text);
                    paragraph.Append(paragraphProperties);
                    paragraph.Append(run);

                    body.Append(paragraph);

                    document.Append(body);
                    wordDocument.MainDocumentPart.Document = document;
                    wordDocument.Close();
                }

                return File(stream.ToArray(), docxMIMEType,
                    "Word Document Basic Example.docx");
            }
        }

        public IActionResult CreateComplexDocument()
        {
            using (var stream = new MemoryStream())
            {
                using (var wordDocument = WordprocessingDocument.Create(stream,
                    WordprocessingDocumentType.Document, true))
                {
                    wordDocument.AddMainDocumentPart();

                    var document = new Document();
                    var body = new Body();

                    var paragraph1 = new Paragraph();
                    var paragraphProperties1 = new ParagraphProperties();
                    var paragraphStyleId1 = new ParagraphStyleId() { Val = "Normal" };
                    var justifyJustification1 = new Justification() { Val = JustificationValues.Both };

                    paragraphProperties1.Append(paragraphStyleId1);
                    paragraphProperties1.Append(justifyJustification1);

                    var run1 = new Run();
                    var text1 = new Text("A normal text, ") { Space = SpaceProcessingModeValues.Preserve };
                    run1.Append(text1);

                    ///////////////////////////

                    var run2 = new Run();
                    var runProperties2 = new RunProperties();
                    runProperties2.Bold = new Bold();

                    var text2 = new Text("now a bold text, ") { Space = SpaceProcessingModeValues.Preserve };
                    run2.Append(runProperties2); // Properties must go first... always!
                    run2.Append(text2);

                    ///////////////////////////

                    var run3 = new Run();
                    var runProperties3 = new RunProperties();
                    runProperties3.Italic = new Italic();

                    var text3 = new Text("now an italic text. ") { Space = SpaceProcessingModeValues.Preserve };
                    run3.Append(runProperties3);
                    run3.Append(text3);

                    ///////////////////////////

                    var run4 = new Run();
                    var runProperties4 = new RunProperties();
                    runProperties4.Italic = new Italic();
                    runProperties4.Underline = new Underline();
                    runProperties4.Bold = new Bold();

                    var text4 = new Text("Yes, you can combine styles, ") { Space = SpaceProcessingModeValues.Preserve };
                    run4.Append(runProperties4);
                    run4.Append(text4);

                    ///////////////////////////

                    var run5 = new Run();
                    var runProperties5 = new RunProperties();
                    runProperties5.Color = new Color() { Val = "FFFF00" };

                    var text5 = new Text("and add some color for your text.") { Space = SpaceProcessingModeValues.Preserve };
                    run5.Append(runProperties5);
                    run5.Append(text5);

                    paragraph1.Append(paragraphProperties1);
                    paragraph1.Append(run1);
                    paragraph1.Append(run2);
                    paragraph1.Append(run3);
                    paragraph1.Append(run4);
                    paragraph1.Append(run5);

                    ///////////////////////////
                    ///////////////////////////

                    var runProperties6 = new RunProperties();
                    runProperties6.RunStyle = new RunStyle() { Val = "Hyperlink" };
                    runProperties6.Color = new Color() { ThemeColor = ThemeColorValues.Hyperlink };

                    var text6 = new Text("Visit my website!") { Space = SpaceProcessingModeValues.Preserve };
                    var run6 = new Run(runProperties6, text6);

                    var url = "https://luisbeltran.mx";
                    var uri = new Uri(url);

                    var hyperlinkRelationship = wordDocument.MainDocumentPart.AddHyperlinkRelationship(uri, true);
                    var id = hyperlinkRelationship.Id;

                    var proofError = new ProofError() { Type = ProofingErrorValues.GrammarStart };

                    var hyperLink = new Hyperlink(proofError, run6)
                    {
                        History = OnOffValue.FromBoolean(true),
                        Id = id
                    };

                    var paragraph2 = new Paragraph();
                    paragraph2.Append(hyperLink);

                    ///////////////////////////

                    body.Append(paragraph1);
                    body.Append(paragraph2);

                    document.Append(body);
                    wordDocument.MainDocumentPart.Document = document;
                    wordDocument.Close();
                }

                return File(stream.ToArray(), docxMIMEType,
                    "Word Document Complex Example.docx");
            }
        }

        public IActionResult AddTable()
        {
            using (var stream = new MemoryStream())
            {
                using (var wordDocument = WordprocessingDocument.Create(stream,
                    WordprocessingDocumentType.Document, true))
                {
                    wordDocument.AddMainDocumentPart();

                    var document = new Document();
                    var body = new Body();

                    var table = new Table();

                    var tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
                    var borderColor = "FF8000";

                    var tableProperties = new TableProperties();
                    var tableBorders = new TableBorders();

                    var topBorder = new TopBorder();
                    topBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                    topBorder.Color = borderColor;

                    var bottomBorder = new BottomBorder();
                    bottomBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                    bottomBorder.Color = borderColor;

                    var rightBorder = new RightBorder();
                    rightBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                    rightBorder.Color = borderColor;

                    var leftBorder = new LeftBorder();
                    leftBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                    leftBorder.Color = borderColor;

                    var insideHorizontalBorder = new InsideHorizontalBorder();
                    insideHorizontalBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                    insideHorizontalBorder.Color = borderColor;

                    var insideVerticalBorder = new InsideVerticalBorder();
                    insideVerticalBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
                    insideVerticalBorder.Color = borderColor;

                    tableBorders.AppendChild(topBorder);
                    tableBorders.AppendChild(bottomBorder);
                    tableBorders.AppendChild(rightBorder);
                    tableBorders.AppendChild(leftBorder);
                    tableBorders.AppendChild(insideHorizontalBorder);
                    tableBorders.AppendChild(insideVerticalBorder);

                    tableProperties.Append(tableWidth);
                    tableProperties.AppendChild(tableBorders);

                    table.AppendChild(tableProperties);

                    //////////

                    var row1 = new TableRow();
                    var cell11 = new TableCell();
                    var paragraph11 = new Paragraph(new Run(new Text("Luis")));

                    cell11.Append(paragraph11);
                    row1.Append(cell11);
                    ///// 

                    var cell12 = new TableCell();
                    var paragraph12 = new Paragraph();
                    var run12 = new Run();
                    var runProperties12 = new RunProperties();
                    runProperties12.Bold = new Bold();

                    run12.Append(runProperties12);
                    run12.Append(new Text("400"));
                    paragraph12.Append(run12);
                    cell12.Append(paragraph12);
                    row1.Append(cell12);
                    /////

                    table.Append(row1);

                    var random = new Random();

                    for (int i = 1; i < 5; i++)
                    {
                        var row = new TableRow();

                        var cell1 = new TableCell();
                        var paragraph = new Paragraph(new Run(new Text($"Employee {i}")));
                        cell1.Append(paragraph);

                        var cell2 = new TableCell();
                        var paragraph2 = new Paragraph();
                        var paragraphProperties2 = new ParagraphProperties();
                        paragraphProperties2.Justification = new Justification() { Val = JustificationValues.Center };
                        paragraph2.Append(paragraphProperties2);
                        paragraph2.Append(new Run(new Text(random.Next(100, 500).ToString())));
                        cell2.Append(paragraph2);

                        row.Append(cell1);
                        row.Append(cell2);

                        table.Append(row);
                    }

                    body.Append(table);

                    document.Append(body);
                    wordDocument.MainDocumentPart.Document = document;
                    wordDocument.Close();
                }

                return File(stream.ToArray(), docxMIMEType,
                    "Word Document Table Example.docx");
            }
        }

        public IActionResult AddList()
        {
            using (var stream = new MemoryStream())
            {
                using (var wordDocument = WordprocessingDocument.Create(stream,
                    WordprocessingDocumentType.Document, true))
                {
                    wordDocument.AddMainDocumentPart();

                    var document = new Document();
                    var body = new Body();

                    var spacing = new SpacingBetweenLines() { After = "0" };
                    var indentation = new Indentation() { Left = "5", Hanging = "360" };
                    var numberingProperties = new NumberingProperties(
                        new NumberingLevelReference() { Val = 1 }, //ordered list
                        new NumberingId() { Val = 2 } //ordered list
                    );

                    var paragraphProperties = new ParagraphProperties(numberingProperties, spacing, indentation);
                    paragraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = "ListParagraph" };

                    var paragraph0 = new Paragraph(new Run(new Text("Ordered List.")));

                    var paragraph1 = new Paragraph();
                    paragraph1.ParagraphProperties = new ParagraphProperties(paragraphProperties.OuterXml);
                    paragraph1.Append(new Run(new Text("Soccer")));

                    Paragraph paragraph2 = new Paragraph();
                    paragraph2.ParagraphProperties = new ParagraphProperties(paragraphProperties.OuterXml);
                    paragraph2.Append(new Run(new Text("Basketball")));

                    Paragraph paragraph3 = new Paragraph();
                    paragraph3.ParagraphProperties = new ParagraphProperties(paragraphProperties.OuterXml);
                    paragraph3.Append(new Run(new Text("Tennis")));

                    body.Append(paragraph0);
                    body.Append(paragraph1);
                    body.Append(paragraph2);
                    body.Append(paragraph3);

                    /////////////////////////////

                    NumberingDefinitionsPart numberingPart = wordDocument.MainDocumentPart.NumberingDefinitionsPart;
                    if (numberingPart == null)
                        numberingPart = wordDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>("NumberingDefinitionsPart001");

                    //var numberingPart = wordDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>("NumberingDefinitionsPart001"); //unique ID

                    var numbering =
                      new Numbering(
                        new AbstractNum(
                          new Level(
                            new NumberingFormat() { Val = NumberFormatValues.Bullet },
                            new LevelText() { Val = "·" } //♦
                          )
                          { LevelIndex = 0 }
                        )
                        { AbstractNumberId = 1 },
                        new NumberingInstance(
                          new AbstractNumId() { Val = 1 }
                        )
                        { NumberID = 1 });

                    numbering.Save(numberingPart);

                    var spacing2 = new SpacingBetweenLines() { After = "5" };
                    var indentation2 = new Indentation() { Left = "5", Hanging = "360" };
                    var numberingProperties2 = new NumberingProperties(
                        new NumberingLevelReference() { Val = 0 }, //unordered list
                        new NumberingId() { Val = 1 } //unordered list
                    );

                    var paragraphProperties2 = new ParagraphProperties(numberingProperties2, spacing2, indentation2);
                    paragraphProperties2.ParagraphStyleId = new ParagraphStyleId() { Val = "ListParagraph" };

                    var paragraph20 = new Paragraph(new Run(new Text("Unordered List.")));
                    body.Append(paragraph20);

                    var countryList = new List<string>() { "Mexico", "Czech Republic", "Nicaragua", "Italy" };

                    foreach (var item in countryList)
                    {
                        var paragraph2n = new Paragraph();
                        paragraph2n.ParagraphProperties = new ParagraphProperties(paragraphProperties2.OuterXml);
                        paragraph2n.Append(new Run(new Text(item)));

                        body.Append(paragraph2n);
                    }

                    document.Append(body);
                    wordDocument.MainDocumentPart.Document = document;
                    wordDocument.Close();
                }

                return File(stream.ToArray(), docxMIMEType,
                    "Word Document List Example.docx");
            }
        }

        public IActionResult OpenEdit()
        {
            var fileName = "Word Document Basic Example.docx";
            var filePath = $@"C:\Users\tony_\Downloads\{fileName}";

            using (var wordDocument = WordprocessingDocument.Open(filePath, true))
            {
                var body = wordDocument.MainDocumentPart.Document.Body;

                var paragraph = body.AppendChild(new Paragraph());
                var run = paragraph.AppendChild(new Run());
                run.AppendChild(new Text("This is a new text"));
                wordDocument.Close();
            }

            return File(System.IO.File.ReadAllBytes(filePath), docxMIMEType, fileName);
        }

        public IActionResult CreateFromTemplate()
        {
            var fileName = "Form Template.docx";
            var filePath = $@"C:\Users\tony_\Downloads\{fileName}";

            var name = "Luis";
            var country = "Czech Republic";

            using (var wordDocument = WordprocessingDocument.Open(filePath, true))
            {
                var body = wordDocument.MainDocumentPart.Document.Body;
                var fields = body.Descendants<FormFieldData>();

                foreach (var field in fields)
                {
                    var formField = (FormFieldName)field.FirstChild;
                    var fieldName = formField.Val.InnerText;

                    switch (fieldName)
                    {
                        case "Name":
                            UpdateFormField(field, name);
                            break;
                        case "Country":
                            UpdateFormField(field, country);
                            break;
                        default:
                            break;
                    }
                }

                wordDocument.Close();
            }

            return File(System.IO.File.ReadAllBytes(filePath), docxMIMEType, $"{name} Info.docx");
        }

        private void UpdateFormField(FormFieldData field, string value)
        {
            var text = field.Descendants<TextInput>().First();
            WordHelpers.SetFormFieldValue(text, value);
        }
    }
}
