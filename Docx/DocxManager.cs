using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using System.Net;
using Newtonsoft.Json;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DE = DocumentFormat.OpenXml.Office.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
using TopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder;
using BottomBorder = DocumentFormat.OpenXml.Wordprocessing.BottomBorder;
using RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder;
using LeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder;
using InsideHorizontalBorder = DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder;
using InsideVerticalBorder = DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;
using DocumentFormat.OpenXml.Presentation;
using TextBody = DocumentFormat.OpenXml.Presentation.TextBody;
using Shape = DocumentFormat.OpenXml.Presentation.Shape;
using Z.Expressions;
using NonVisualGroupShapeProperties = DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties;
using NonVisualDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties;
using NonVisualGroupShapeDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties;
using NonVisualShapeProperties = DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties;
using NonVisualShapeDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties;
using ShapeProperties = DocumentFormat.OpenXml.Presentation.ShapeProperties;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using Picture = DocumentFormat.OpenXml.Drawing.Picture;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using NonVisualPictureProperties = DocumentFormat.OpenXml.Presentation.NonVisualPictureProperties;
using NonVisualGraphicFrameDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualGraphicFrameDrawingProperties;
using NonVisualPictureDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualPictureDrawingProperties;
using BlipFill = DocumentFormat.OpenXml.Presentation.BlipFill;
using NonVisualGraphicFrameProperties = DocumentFormat.OpenXml.Presentation.NonVisualGraphicFrameProperties;
using GraphicFrame = DocumentFormat.OpenXml.Presentation.GraphicFrame;
using System.Reflection.Metadata;
using static System.Net.Mime.MediaTypeNames;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using PresentationDocument = DocumentFormat.OpenXml.Packaging.PresentationDocument;
using Slide = DocumentFormat.OpenXml.Presentation.Slide;
using System.Linq;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Spreadsheet;
using ETable = DocumentFormat.OpenXml.Spreadsheet.Table;
using E = DocumentFormat.OpenXml.Spreadsheet;
using Bold = DocumentFormat.OpenXml.Wordprocessing.Bold;
using Italic = DocumentFormat.OpenXml.Wordprocessing.Italic;
using OfficeOpenXml.Table;
using OfficeOpenXml;
using System.Drawing.Imaging;
using System.Drawing;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Transform = DocumentFormat.OpenXml.Presentation.Transform;
using GroupShapeProperties = DocumentFormat.OpenXml.Presentation.GroupShapeProperties;
using OfficeOpenXml.Drawing;
using EP = OfficeOpenXml;
using NCalc;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Math;
using JustificationValues = DocumentFormat.OpenXml.Wordprocessing.JustificationValues;
using Justification = DocumentFormat.OpenXml.Wordprocessing.Justification;
using Microsoft.Office.Interop.PowerPoint;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using Row = DocumentFormat.OpenXml.Spreadsheet.Row;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using I = Microsoft.Office.Interop.PowerPoint;
using MOC = Microsoft.Office.Core;
using System.Runtime.InteropServices;


namespace Docx
{
    public class DocxManager
    {
        private readonly ILogger<DocxManager> _logger;

        public DocxManager(ILogger<DocxManager> logger)
        {
            _logger = logger;
        }

        class BulletedListItem
        {
            public string OperatingSystem { get; set; }
            public string Price { get; set; }
            public string Reference { get; set; }
        }

        [Function("WordFunction")]
        public async Task<HttpResponseData> WordFunction([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
        {
            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/json; charset=utf-8");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // For non-commercial use

            {
                try
                {
                    _logger.LogInformation("C# HTTP trigger function 'pdfScan' processed a request.");
                    string requestBody = await new StreamReader(req.Body).ReadToEndAsync();

                    string sourcefilePath = "C:\\Users\\Aux-143\\Desktop\\sample\\sample doc\\Testing Document.docx";

                    string filePath = "C:\\Users\\Aux-143\\Desktop\\sample\\sample doc\\tested.docx";

                    string efilePath = "C:\\Users\\Aux-143\\Desktop\\sample\\sample doc\\Table.xlsx";


                    CopyFile(sourcefilePath, filePath);

                    dynamic jsonData = JsonConvert.DeserializeObject(requestBody);
                    dynamic jsonObject = jsonData[0];
                    dynamic tableData = jsonObject.tableData;
                    dynamic bulletedList = jsonObject.bulletedListItems;


                    string headerValue = jsonObject.headerValue;
                    string testTag1Value = jsonObject.testTag1Value;
                    string testTag2Value = jsonObject.testTag2Value;
                    string testTag3Value = jsonObject.testTag3Value;
                    string testTag6Value = jsonObject.testTag6Value;

                    string base64image1 = jsonObject.base64image1;
                    string base64image2 = jsonObject.base64image2;


                    List<string[]> dataList = new List<string[]>();
                    foreach (var person in tableData)
                    {
                        string[] rowData = new string[]
                        {
                            person.ID,
                            person.Name,
                            person.Age,
                            person.PhoneNumber,
                            person.City
                        };
                        dataList.Add(rowData);
                    }

                    List<BulletedListItem> bulletedListItems = new List<BulletedListItem>();
                    foreach (var item in bulletedList)
                    {
                        // Create a new BulletedListItem instance
                        BulletedListItem bulletedListItem = new BulletedListItem();

                        // Assign values from the dynamic object to the properties of BulletedListItem
                        bulletedListItem.OperatingSystem = item.OperatingSystem;
                        bulletedListItem.Price = item.Price;
                        bulletedListItem.Reference = item.Reference;

                        // Add the bulletedListItem to the dataList
                        bulletedListItems.Add(bulletedListItem);
                    }

                    Table table = new Table(
                         new TableProperties(
                            new TableBorders(
                                new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                                new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                                new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                                new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                                new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 }
                            )
                        ),
                             new TableRow(
                                new TableCell(new Paragraph(new Run(new Text("Product")))),
                                new TableCell(new Paragraph(new Run(new Text("Price")))),
                                new TableCell(new Paragraph(new Run(new Text("Quantity"))))
                            ),
                             new TableRow(
                                new TableCell(new Paragraph(new Run(new Text("Apples")))),
                                new TableCell(new Paragraph(new Run(new Text("$1.50")))),
                                new TableCell(new Paragraph(new Run(new Text("10"))))
                            ),
                             new TableRow(
                                new TableCell(new Paragraph(new Run(new Text("Bananas")))),
                                new TableCell(new Paragraph(new Run(new Text("$0.75")))),
                                new TableCell(new Paragraph(new Run(new Text("15"))))
                            )
                     );

                    //word file manipulation

                    ReplaceTags(filePath, "#HeaderTag", headerValue);
                    ReplaceTags(filePath, "#testTag1", testTag1Value);
                    ReplaceTags(filePath, "#testTag2", testTag2Value);
                    ReplaceTags(filePath, "#testTag3", testTag3Value);
                    ReplaceTags(filePath, "#tagParagraph", testTag6Value);

                    ReplaceTagsByImages(filePath, "#testTag4", base64image1);
                    ReplaceTagsByImages(filePath, "#testTag5", base64image2);

                    InsertTable(filePath, dataList);

                    GenerateBulletedList(filePath, bulletedListItems);

                    ReplaceTagsByTables(filePath, "tableTag", table);

                    string StyleName = "Heading2";
                    string Tag = "#Header2TagStyle";

                    //styling
                    AddTag1Style(filePath, testTag1Value);
                    ApplyParagraphAlignment(filePath, testTag6Value, JustificationValues.Center);
                    AddHeadingStyle(filePath, StyleName, Tag);

                    //apply styles to the table
                    int tableNumber = 1;
                    ApplyStyleToFullTable(filePath, tableNumber);
                    //apply styling to a specific cell
                    int targetRow = 2;
                    int targetColumn = 3;
                    ApplySubtitleStyleToCell(filePath, targetRow, targetColumn);

                    //condition checking
                    EvaluateCondition(filePath, testTag1Value, testTag2Value, testTag3Value);                 

                }

                catch (Exception ex)
                {
                    _logger.LogError($"Error processing the request: {ex.Message}");
                    response.WriteString($"Error processing the request: {ex.Message}");
                }
            }
            return response;

        }

        [Function("PPTFunction")]
        public async Task<HttpResponseData> PPTFunction([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
        {
            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/json; charset=utf-8");

            {
                try
                {
                    _logger.LogInformation("C# HTTP trigger function 'pdfScan' processed a request.");
                    string requestBody = await new StreamReader(req.Body).ReadToEndAsync();

                    string sourcePPTfilePath = "C:\\Users\\Aux-143\\Desktop\\sample\\sample doc\\TestSample.pptx";

                    string pfilePath = "C:\\Users\\Aux-143\\Desktop\\sample\\sample doc\\TestedSample.pptx";

                    string efilePath = "C:\\Users\\Aux-143\\Desktop\\sample\\sample doc\\Table.xlsx";


                    string outputImagePath = "C:\\Users\\Aux-143\\Desktop\\sample\\sample doc\\output.png";

                    CopyFile(sourcePPTfilePath, pfilePath);

                    dynamic jsonData = JsonConvert.DeserializeObject(requestBody);
                    dynamic jsonObject = jsonData[0];
                    dynamic tableData = jsonObject.tableData;
                    dynamic bulletedList = jsonObject.bulletedListItems;


                    string headerValue = jsonObject.headerValue;
                    string testTag1Value = jsonObject.testTag1Value;
                    string testTag2Value = jsonObject.testTag2Value;
                    string testTag3Value = jsonObject.testTag3Value;
                    string testTag6Value = jsonObject.testTag6Value;

                    string base64image1 = jsonObject.base64image1;
                    string base64image2 = jsonObject.base64image2;


                    List<string[]> dataList = new List<string[]>();
                    foreach (var person in tableData)
                    {
                        string[] rowData = new string[]
                        {
                            person.ID,
                            person.Name,
                            person.Age,
                            person.PhoneNumber,
                            person.City
                        };
                        dataList.Add(rowData);
                    }

                    List<BulletedListItem> bulletedListItems = new List<BulletedListItem>();
                    foreach (var item in bulletedList)
                    {
                        // Create a new BulletedListItem instance
                        BulletedListItem bulletedListItem = new BulletedListItem();

                        // Assign values from the dynamic object to the properties of BulletedListItem
                        bulletedListItem.OperatingSystem = item.OperatingSystem;
                        bulletedListItem.Price = item.Price;
                        bulletedListItem.Reference = item.Reference;

                        // Add the bulletedListItem to the dataList
                        bulletedListItems.Add(bulletedListItem);
                    }

                    Table table = new Table(
                         new TableProperties(
                            new TableBorders(
                                new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                                new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                                new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                                new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 },
                                new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 6 }
                            )
                        ),
                             new TableRow(
                                new TableCell(new Paragraph(new Run(new Text("Product")))),
                                new TableCell(new Paragraph(new Run(new Text("Price")))),
                                new TableCell(new Paragraph(new Run(new Text("Quantity"))))
                            ),
                             new TableRow(
                                new TableCell(new Paragraph(new Run(new Text("Apples")))),
                                new TableCell(new Paragraph(new Run(new Text("$1.50")))),
                                new TableCell(new Paragraph(new Run(new Text("10"))))
                            ),
                             new TableRow(
                                new TableCell(new Paragraph(new Run(new Text("Bananas")))),
                                new TableCell(new Paragraph(new Run(new Text("$0.75")))),
                                new TableCell(new Paragraph(new Run(new Text("15"))))
                            )
                     );


                    AddNewSlide(pfilePath);

                    ReplaceTagsByWordsInSlide(pfilePath, "#testTag1", testTag1Value);
                    ReplaceTagsByWordsInSlide(pfilePath, "#testTag2", testTag2Value);
                    ReplaceTagsByWordsInSlide(pfilePath, "#testTag3", testTag3Value);


                    //ExtractTableFromSlide(pfilePath, 0);

                    EvaluateConditionInSlide(pfilePath, testTag1Value, testTag2Value, testTag3Value);

                    ////CreateExcelTableAndInsertIntoPowerPoint(pfilePath, efilePath);

                    GenerateBulletedListInPPT(pfilePath, bulletedListItems);

                    ReplaceImageByImages(pfilePath, base64image2, 3);

                }

                catch (Exception ex)
                {
                    _logger.LogError($"Error processing the request: {ex.Message}");
                    response.WriteString($"Error processing the request: {ex.Message}");
                }
            }
            return response;

        }

        static void CopyFile(string sourceFilePath, string destinationFilePath)
        {
            if (File.Exists(sourceFilePath))
            {
                // Create a new package for the destination file
                using (Package sourcePackage = Package.Open(sourceFilePath, FileMode.Open, FileAccess.Read))
                using (Package destinationPackage = Package.Open(destinationFilePath, FileMode.Create))
                {
                    // Copy each part from the source to the destination
                    foreach (PackagePart part in sourcePackage.GetParts())
                    {
                        // Create the part URI for the destination package
                        Uri partUri = PackUriHelper.CreatePartUri(new Uri(part.Uri.OriginalString, UriKind.Relative));

                        // Create the part in the destination package
                        PackagePart destinationPart = destinationPackage.CreatePart(partUri, part.ContentType, CompressionOption.Normal);

                        // Copy the content from the source part to the destination part
                        using (Stream sourceStream = part.GetStream())
                        using (Stream destinationStream = destinationPart.GetStream())
                        {
                            sourceStream.CopyTo(destinationStream);
                        }
                    }
                }

                Console.WriteLine($"File '{sourceFilePath}' copied successfully.");
            }
            else
            {
                Console.WriteLine($"Source file '{sourceFilePath}' does not exist.");
            }
        }

        private void ReplaceTags(string filePath, string tag, string replacement)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                var textElements = doc.MainDocumentPart.Document.Body.Descendants<Text>()
                .Where(t => t.Text.Contains(tag));

                foreach (var text in textElements)
                {
                    text.Text = text.Text.Replace(tag, replacement);
                }
                doc.MainDocumentPart.Document.Save();
            }
        }

        public void ConditionChecking(string filePath, string tagToCheck1, string tagToCheck2, string tagToUpdate)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                var textElements = doc.MainDocumentPart.Document.Body.Descendants<Text>()
                .Where(t => t.Text.Contains(tagToUpdate));

                foreach (var text in textElements)
                {
                    //var parent = text.Parent;

                    if (tagToCheck1 != null && tagToCheck2 != null &&
                   tagToCheck1 == "No" && tagToCheck2 == "Yes")
                    {

                        // Remove the parent element (e.g., paragraph) containing the tag to update
                        text.Remove();

                    }
                }
                doc.MainDocumentPart.Document.Save();
            }

        }

        public void EvaluateCondition(string filePath, string tag1Value, string tag2Value, string tag3Value)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                var textElements = doc.MainDocumentPart.Document.Body.Descendants<Text>()
                    .Where(t => t.Text.Contains(tag3Value));

                foreach (var condition in doc.MainDocumentPart.Document.Body.Descendants<Text>()
                    .Where(t => t.Text != null && t.Text.Contains("==")))
                {
                    foreach (var text in textElements)
                    {
                        string conditionText = condition.Text.Replace("\"\"", "\"") // Remove extra double quotes
                        .Replace("$tag1", tag1Value) // Replace tag placeholders with actual values
                        .Replace("$tag2", tag2Value);

                        bool isShowable = Eval.Execute<bool>(conditionText);

                        if (!isShowable)
                        {
                            text.Remove();
                        }
                    }
                }
            }
        }

        static void AddTag1Style(string filePath, string tagToStyle)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                var body = doc.MainDocumentPart.Document.Body;

                // Apply the style to paragraphs containing the specified tag
                foreach (var paragraph in body.Elements<Paragraph>())
                {
                    if (paragraph.InnerText.Contains(tagToStyle))
                    {
                        // Apply formatting properties directly to the runs within the paragraph
                        foreach (var run in paragraph.Elements<Run>())
                        {
                            // Add bold and italic properties to the run
                            var runProperties = run.RunProperties ?? (run.RunProperties = new RunProperties());
                            runProperties.AppendChild(new Bold());
                            runProperties.AppendChild(new Italic());

                            // Add font color red
                            Color color = new Color() { Val = "FF0000" }; // Red color
                            runProperties.AppendChild(color);
                        }
                    }
                }
            }
        }

        static void ApplyParagraphAlignment(string filePath, string tagToStyle, JustificationValues alignment)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                var body = doc.MainDocumentPart.Document.Body;

                // Apply the alignment to paragraphs containing the specified tag
                foreach (var paragraph in body.Elements<Paragraph>())
                {
                    if (paragraph.InnerText.Contains(tagToStyle))
                    {
                        // Check if ParagraphProperties exist, if not, create them
                        if (paragraph.Elements<ParagraphProperties>().Count() == 0)
                        {
                            paragraph.PrependChild(new ParagraphProperties());
                        }

                        // Apply the alignment
                        paragraph.GetFirstChild<ParagraphProperties>().Append(new Justification() { Val = alignment });
                    }
                }
            }
        }

        static void ReplaceTagsByImages(string filePath, string tag, string base64Image)
        {
            byte[] imageBytes = Convert.FromBase64String(base64Image);

            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
                {
                    var textElements = doc.MainDocumentPart.Document.Body.Descendants<Text>()
                                            .Where(t => t.Text.Contains(tag));

                    foreach (var text in textElements.ToList())
                    {
                        // Create the Drawing element
                        var drawing = CreateDrawingElementinWord(doc.MainDocumentPart, ms);

                        // Create a new Run with the Drawing element
                        var run = new Run(drawing);

                        // Replace the text with the new Run containing the image
                        var parent = text.Parent;

                        parent.ReplaceChild(run, text);
                    }

                    doc.MainDocumentPart.Document.Save();
                }
            }
        }

        static Drawing CreateDrawingElementinWord(MainDocumentPart mainPart, MemoryStream stream)
        {
            // Create a new image part
            var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
            imagePart.FeedData(stream);

            // Generate the relationship ID for the image
            string relationshipId = mainPart.GetIdOfPart(imagePart);

            // Create the Drawing element
            var drawing = new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = 990000L, Cy = 792000L },
                    new DW.EffectExtent()
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DW.DocProperties()
                    {
                        Id = (UInt32Value)1U,
                        Name = "Picture 1"
                    },
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties()
                                    {
                                        Id = (UInt32Value)0U,
                                        Name = "New Bitmap Image.png"
                                    },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip()
                                    {
                                        Embed = relationshipId,
                                        CompressionState = A.BlipCompressionValues.Print
                                    },
                                    new A.Stretch(
                                        new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset() { X = 0L, Y = 0L },
                                        new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                    new A.PresetGeometry(
                                        new A.AdjustValueList()
                                    )
                                    { Preset = A.ShapeTypeValues.Rectangle }))
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                )
                {
                    DistanceFromTop = (UInt32Value)0U,
                    DistanceFromBottom = (UInt32Value)0U,
                    DistanceFromLeft = (UInt32Value)0U,
                    DistanceFromRight = (UInt32Value)0U,
                    EditId = "50D07946"
                });

            return drawing;
        }

        static void InsertTable(string filePath, List<string[]> data)
        {
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
                {
                    var mainPart = doc.MainDocumentPart;
                    var existingTable = mainPart.Document.Body.Elements<Table>().FirstOrDefault();
                    if (existingTable == null)
                    {
                        Console.WriteLine("No existing table found in the document.");
                        return;
                    }

                    // Get the headers from the first row of the existing table
                    var headers = existingTable.Elements<TableRow>().First().Elements<TableCell>().Select(cell => cell.InnerText).ToArray();

                    // Add rows with data to the existing table.
                    foreach (var rowData in data)
                    {
                        // Create a new row
                        var row = new TableRow();

                        // Add cells with data to the row
                        for (int i = 0; i < Math.Min(headers.Length, rowData.Length); i++)
                        {
                            var cell = new DocumentFormat.OpenXml.Wordprocessing.TableCell(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run(new Text(rowData[i]))));
                            row.Append(cell);
                        }

                        // Append the row to the existing table
                        existingTable.Append(row);
                    }
                    mainPart.Document.Save();
                }
            }
        }

        static void GenerateBulletedList(string filePath, List<BulletedListItem> dataList)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                // Get the main document part.
                var mainPart = doc.MainDocumentPart;

                // Create a list to store the new paragraphs
                List<Paragraph> newParagraphs = new List<Paragraph>();

                // Find and replace tags in the bulleted list
                foreach (var paragraph in mainPart.Document.Body.Elements<Paragraph>().ToList())
                {
                    if (paragraph.InnerText.Contains("#operatingSystem") &&
                    paragraph.InnerText.Contains("#Price") &&
                    paragraph.InnerText.Contains("#Reference"))
                    {
                        // Remove the existing paragraph
                        paragraph.Remove();

                        // Insert a new paragraph for each item in the list
                        foreach (var item in dataList)
                        {
                            var newParagraph = new Paragraph(
                            new ParagraphProperties(
                            new NumberingProperties(
                            new NumberingLevelReference() { Val = 0 },
                            new NumberingId() { Val = 1 }
                            )
                            ),
                            new Run(
                            new Text($"{item.OperatingSystem} {item.Price} {item.Reference}")
                            )
                            );

                            // Add the new paragraph to the list
                            newParagraphs.Add(newParagraph);
                        }
                    }
                }

                // Insert the new paragraphs at the end of the document
                foreach (var newParagraph in newParagraphs)
                {
                    mainPart.Document.Body.AppendChild(newParagraph);
                }

                // Save the changes.
                mainPart.Document.Save();
            }
        }

        static void ReplaceTagsByTables(string filePath, string tag, Table table)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                // Find all text elements containing the tag
                var textElements = doc.MainDocumentPart.Document.Body.Descendants<Text>()
                    .Where(t => t.Text.Contains(tag)).ToList();

                foreach (var text in textElements)
                {
                    // Create a table

                    // Get the parent of the text element
                    var parent = text.Parent;

                    // Replace the text element with the table
                    parent.ReplaceChild(table, text);
                }

                doc.MainDocumentPart.Document.Save();
            }
        }

        static void ApplySubtitleStyleToCell(string filePath, int targetRow, int targetColumn)
        {
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
                {
                    Table firstTable = doc.MainDocumentPart.Document.Body.Elements<Table>().FirstOrDefault();

                    if (firstTable != null)
                    {
                        int currentRow = 0;
                        foreach (TableRow row in firstTable.Elements<TableRow>())
                        {
                            currentRow++;

                            if (currentRow != targetRow)
                                continue;

                            int currentColumn = 0;
                            foreach (TableCell cell in row.Elements<TableCell>())
                            {
                                currentColumn++;

                                if (currentColumn != targetColumn)
                                    continue;

                                foreach (Paragraph paragraph in cell.Elements<Paragraph>())
                                {
                                    ParagraphProperties headingProps = new ParagraphProperties(new ParagraphStyleId() { Val = "Title" });
                                    paragraph.PrependChild(headingProps.CloneNode(true));
                                }
                                break;
                            }
                            break;
                        }
                        doc.MainDocumentPart.Document.Save();
                    }
                    else
                    {
                        Console.WriteLine("No table found in the document.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        static void ApplyStyleToFullTable(string filePath, int tableNumber)
        {
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
                {
                    var tables = doc.MainDocumentPart.Document.Body.Elements<Table>().ToList();

                    if (tables.Count >= tableNumber)
                    {
                        Table targetTable = tables[tableNumber - 1];

                        foreach (TableRow row in targetTable.Elements<TableRow>())
                        {
                            foreach (TableCell cell in row.Elements<TableCell>())
                            {
                                foreach (Paragraph paragraph in cell.Elements<Paragraph>())
                                {
                                    ParagraphProperties headingProps = new ParagraphProperties(new ParagraphStyleId() { Val = "Subtitle" });
                                    paragraph.PrependChild(headingProps.CloneNode(true));
                                }
                            }
                        }
                        doc.MainDocumentPart.Document.Save();
                    }
                    else
                    {
                        Console.WriteLine("Table number is out of range.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        static void AddHeadingStyle(string filePath, string StyleName, string tagToStyle)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                var body = doc.MainDocumentPart.Document.Body;
                ParagraphProperties Props = new ParagraphProperties(new ParagraphStyleId() { Val = StyleName });

                foreach (var paragraph in body.Elements<Paragraph>())
                {
                    if (paragraph.InnerText.Contains(tagToStyle))
                    {
                        paragraph.PrependChild(Props.CloneNode(true));
                    }
                }
            }
        }

        //power point methods
        static void AddNewSlide(string presentationFile)
        {
            // Open the source document as read/write. 
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
            {
                PresentationPart? presentationPart = presentationDocument.PresentationPart;

                // Verify that the presentation is not empty.
                if (presentationPart is null)
                {
                    throw new InvalidOperationException("The presentation document is empty.");
                }

                // Declare and instantiate a new slide.
                Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
                uint drawingObjectId = 1;

                // Construct the slide content.            
                // Specify the non-visual properties of the new slide.
                CommonSlideData commonSlideData = slide.CommonSlideData ?? slide.AppendChild(new CommonSlideData());
                ShapeTree shapeTree = commonSlideData.ShapeTree ?? commonSlideData.AppendChild(new ShapeTree());
                NonVisualGroupShapeProperties nonVisualProperties = shapeTree.AppendChild(new NonVisualGroupShapeProperties());
                nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
                nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
                nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

                // Specify the group shape properties of the new slide.
                shapeTree.AppendChild(new GroupShapeProperties());

                // Declare and instantiate the title shape of the new slide.
                Shape titleShape = shapeTree.AppendChild(new Shape());
                drawingObjectId++;

                // Specify the required shape properties for the title shape. 
                titleShape.NonVisualShapeProperties = new NonVisualShapeProperties
                    (new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "" },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
                titleShape.ShapeProperties = new ShapeProperties();

                // Declare and instantiate the body shape of the new slide.
                Shape bodyShape = shapeTree.AppendChild(new Shape());
                drawingObjectId++;

                // Specify the required shape properties for the body shape.
                bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
                        new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                        new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
                bodyShape.ShapeProperties = new ShapeProperties();

                // Create the slide part for the new slide.
                SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

                // Save the new slide part.
                slide.Save(slidePart);

                // Modify the slide ID list in the presentation part.
                // The slide ID list should not be null.
                SlideIdList? slideIdList = presentationPart.Presentation.SlideIdList;

                // Find the highest slide ID in the current list.
                uint maxSlideId = 1;
                foreach (SlideId slideId in slideIdList.ChildElements)
                {
                    if (slideId.Id is not null && slideId.Id > maxSlideId)
                    {
                        maxSlideId = slideId.Id;
                    }
                }

                maxSlideId++;

                // Get the ID of the last slide.
                SlideId lastSlideId = slideIdList.LastChild as SlideId;
                SlidePart lastSlidePart;

                if (lastSlideId is not null && lastSlideId.RelationshipId is not null)
                {
                    lastSlidePart = (SlidePart)presentationPart.GetPartById(lastSlideId.RelationshipId);
                }
                else
                {
                    throw new InvalidOperationException("The last slide ID or its relationship ID is null.");
                }

                // Use the same slide layout as that of the last slide.
                if (lastSlidePart.SlideLayoutPart is not null)
                {
                    slidePart.AddPart(lastSlidePart.SlideLayoutPart);
                }

                // Insert the new slide into the slide list after the last slide.
                SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), lastSlideId);
                newSlideId.Id = maxSlideId;
                newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

                // Save the modified presentation.
                presentationPart.Presentation.Save();
            }
        }

        static void ReplaceTagsByWordsInSlide(string filePath, string tag, string replacement)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
            {
                // Access the main presentation part
                PresentationPart presentationPart = presentationDocument.PresentationPart;

                // Iterate through each slide part in the presentation
                foreach (SlidePart slidePart in presentationPart.SlideParts)
                {
                    // Load the slide XML content
                    Slide slide = slidePart.Slide;

                    // Iterate through each shape in the slide
                    foreach (var shape in slide.Descendants<Shape>())
                    {
                        // Check if the shape contains a text body
                        var textBody = shape.Descendants<TextBody>().FirstOrDefault();
                        if (textBody != null)
                        {
                            // Iterate through each paragraph in the text body
                            foreach (var paragraph in textBody.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
                            {
                                // Iterate through each text element in the paragraph
                                foreach (var textElement in paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                                {
                                    // Check if the text element contains the tag
                                    if (textElement.Text.Contains(tag))
                                    {
                                        // Replace the tag with the replacement text
                                        textElement.Text = textElement.Text.Replace(tag, replacement);
                                    }
                                }
                            }
                        }
                    }

                    // Save the changes to the slide part
                    slidePart.Slide.Save();
                }
            }
        }

        public static void ExtractTableFromSlide(string pptxFilePath, int slideIndex)
        {
            using (PresentationDocument presentationDoc = PresentationDocument.Open(pptxFilePath, false))
            {
                PresentationPart presentationPart = presentationDoc.PresentationPart;
                SlidePart slidePart = presentationPart.SlideParts.ElementAt(slideIndex);

                // Find the table shape in the slide
                GraphicFrame tableGraphicFrame = slidePart.Slide.Descendants<GraphicFrame>()
                    .FirstOrDefault(gf => gf.Descendants<A.Table>().Any());

                if (tableGraphicFrame != null)
                {
                    // Get the table element
                    A.Table table = tableGraphicFrame.Descendants<A.Table>().First();

                    // Extract the content of the table
                    foreach (A.TableRow row in table.Descendants<A.TableRow>())
                    {
                        foreach (A.TableCell cell in row.Descendants<A.TableCell>())
                        {
                            foreach (DocumentFormat.OpenXml.Drawing.Paragraph paragraph in cell.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
                            {
                                foreach (DocumentFormat.OpenXml.Drawing.Run run in paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Run>())
                                {
                                    string text = run.Descendants<DocumentFormat.OpenXml.Drawing.Text>().First().InnerText;
                                }
                            }
                        }
                    }
                }
            }
        }

        static void CreateExcelTableAndInsertIntoPowerPoint(string pfilePath, string efilePath)
        {
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("TableSheet");

                // Create sample data for the table.
                string[] headers = { "Header1", "Header2", "Header3" };
                object[,] data = {
            { "Data1", "Data2", "Data3" },
            { "Data4", "Data5", "Data6" },
            { "Data7", "Data8", "Data9" }
        };

                // Define the table range.
                int startRow = 1;
                int startColumn = 1;
                int endRow = startRow + data.GetLength(0) - 1;
                int endColumn = startColumn + data.GetLength(1) - 1;

                // Populate the worksheet with data.
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startColumn; col <= endColumn; col++)
                    {
                        worksheet.Cells[row, col].Value = data[row - startRow, col - startColumn];
                    }
                }

                // Save the Excel package to a file.
                excelPackage.SaveAs(new System.IO.FileInfo(efilePath));
            }

            // Load the PowerPoint presentation document.
            using (PresentationDocument presentation = PresentationDocument.Open(pfilePath, true))
            {
                // Get the presentation slide part.
                SlidePart slidePart = presentation.PresentationPart.SlideParts.ElementAt(0);

                if (slidePart != null)
                {
                    // Create a new shape to hold the table.
                    Shape shape = new Shape();
                    NonVisualShapeProperties nonVisualShapeProperties = new NonVisualShapeProperties();
                    NonVisualDrawingProperties nonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "Table" };
                    nonVisualShapeProperties.Append(nonVisualDrawingProperties);
                    ShapeProperties shapeProperties = new ShapeProperties();
                    A.Transform2D transform2D = new A.Transform2D();
                    A.Offset offset = new A.Offset() { X = 100000L, Y = 100000L }; // Adjust the position as needed
                    A.Extents extents = new A.Extents() { Cx = 9144000L, Cy = 4572000L }; // Adjust the size as needed
                    transform2D.Append(offset);
                    transform2D.Append(extents);
                    shapeProperties.Append(transform2D);
                    shape.Append(nonVisualShapeProperties);
                    shape.Append(shapeProperties);

                    // Load the Excel file and copy the table data.
                    using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(efilePath, false))
                    {
                        WorkbookPart workbookPart = spreadsheet.WorkbookPart;
                        WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                        // Get the sheet data from the worksheet.
                        SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                        if (sheetData != null)
                        {
                            // Create a new table element.
                            A.Table table = new A.Table();

                            // Iterate over the rows in the sheet data.
                            foreach (Row row in sheetData.Elements<Row>())
                            {
                                // Create a new table row element.
                                A.TableRow tableRow = new A.TableRow();

                                // Iterate over the cells in the row.
                                foreach (Cell cell in row.Elements<Cell>())
                                {
                                    // Get the cell value.
                                    string cellValue = GetCellValue(cell, workbookPart);

                                    // Create a new table cell element with the cell value.
                                    A.TableCell tableCell = new A.TableCell(new A.Paragraph(new A.Run(new A.Text(cellValue))));
                                    tableRow.Append(tableCell);
                                }

                                // Append the table row to the table.
                                table.Append(tableRow);
                            }

                            // Copy the table data to the shape.
                            GraphicFrame graphicFrame = new GraphicFrame();
                            graphicFrame.Append(shape);
                            graphicFrame.NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties();

                            GraphicData graphicData = new GraphicData();
                            graphicData.Uri = new StringValue("http://schemas.openxmlformats.org/drawingml/2006/table");

                            graphicData.Append(table); // Append the table to the graphic data.

                            graphicFrame.Append(graphicData);

                            // Add the shape to the slide.
                            Slide slide = slidePart.Slide;
                            slide.CommonSlideData.ShapeTree.AppendChild(graphicFrame);
                        }

                        // Save the changes to the presentation.
                        slidePart.Slide.Save();
                    }
                }
            }
        }

        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            string cellValue = string.Empty;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                // Get the shared string item corresponding to the cell value.
                SharedStringTablePart sharedStringPart = workbookPart.SharedStringTablePart;
                if (sharedStringPart != null)
                {
                    SharedStringItem sharedStringItem = sharedStringPart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(cell.CellValue.Text));
                    if (sharedStringItem != null)
                    {
                        cellValue = sharedStringItem.InnerText;
                    }
                }
            }
            else
            {
                cellValue = cell.CellValue.Text;
            }

            return cellValue;
        }

        static void GenerateBulletedListInPPT(string filePath, List<BulletedListItem> dataList)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
            {
                // Get the presentation part
                PresentationPart presentationPart = presentationDocument.PresentationPart;

                foreach (SlidePart slidePart in presentationPart.SlideParts)
                {
                    Slide slide = slidePart.Slide;

                    foreach (var shape in slide.Descendants<Shape>())
                    {
                        // Check if the shape contains a text body
                        var textBody = shape.Descendants<TextBody>().FirstOrDefault();
                        if (textBody != null)
                        {
                            // Iterate through each paragraph in the text body
                            foreach (var paragraph in textBody.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
                            {
                                // Iterate through each text element in the paragraph
                                foreach (var textElement in paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                                {
                                    // Check if the text element contains the tag
                                    if ((textElement.Text.Contains("#operatingSystem")) && (textElement.Text.Contains("#Price")) &&
                                            (textElement.Text.Contains("#Reference")))
                                    {
                                        paragraph.Remove();

                                        // Create a new paragraph for each item in the list
                                        foreach (var item in dataList)
                                        {
                                            // Create a new paragraph
                                            var newParagraph = new DocumentFormat.OpenXml.Drawing.Paragraph();

                                            // Create a new run
                                            var run = new A.Run();

                                            // Create a new text element with the item's data
                                            var text = new A.Text($"{item.OperatingSystem} {item.Price} {item.Reference}");

                                            // Add the text to the run
                                            run.AppendChild(text);

                                            // Create a new paragraph properties with the bullet style
                                            var paragraphProperties = new A.ParagraphProperties();

                                            // Create a new bullet list style
                                            var bulletList = new A.BulletFontText();

                                            // Add the bullet list style to the paragraph properties
                                            paragraphProperties.AppendChild(bulletList);

                                            // Add the run and paragraph properties to the paragraph
                                            newParagraph.AppendChild(paragraphProperties);
                                            newParagraph.AppendChild(run);

                                            // Add the paragraph to the slide's text body
                                            if (textBody != null)
                                            {
                                                textBody.AppendChild(newParagraph);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                }
                // Save the changes
                presentationPart.Presentation.Save();
            }
        }

        static void EvaluateConditionInSlide(string pfilePath, string testTag1Value, string testTag2Value, string testTag3Value)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(pfilePath, true))
            {
                // Get the presentation part
                PresentationPart presentationPart = presentationDocument.PresentationPart;

                foreach (SlidePart slidePart in presentationPart.SlideParts)
                {
                    Slide slide = slidePart.Slide;

                    foreach (var shape in slide.Descendants<Shape>())
                    {
                        // Check if the shape contains a text body
                        var textBody = shape.Descendants<TextBody>().FirstOrDefault();
                        if (textBody != null)
                        {
                            var condition = "";
                            // Iterate through each paragraph in the text body
                            foreach (var paragraphh in textBody.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
                            {
                                // Iterate through each text element in the paragraph
                                foreach (var textElement in paragraphh.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                                {

                                    // Check if the text element contains the tag
                                    if ((textElement.Text.Contains("==")))
                                    {
                                        condition = textElement.Text;
                                    }
                                }
                            }
                            foreach (var paragraph in textBody.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
                            {
                                // Iterate through each text element in the paragraph
                                foreach (var textElement in paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                                {
                                    // Check if the text element contains the tag
                                    if ((textElement.Text.Contains(testTag3Value)))
                                    {
                                        string conditionText = condition.Replace("\"\"", "\"") // Remove extra double quotes
                                            .Replace("$tag1", testTag1Value) // Replace tag placeholders with actual values
                                            .Replace("$tag2", testTag2Value);

                                        bool isShowable = Eval.Execute<bool>(conditionText);

                                        if (!isShowable)
                                        {
                                            // Remove the entire paragraph
                                            paragraph.Remove();
                                        }

                                    }
                                }
                            }

                        }
                    }

                }
                // Save the changes
                presentationPart.Presentation.Save();
            }

        }

        public void ReplaceImageByImages(string filePath, string base64image1, int slideIndex)
        {
            byte[] imageBytes = Convert.FromBase64String(base64image1);

            using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
            {
                PresentationPart presentationPart = presentationDocument.PresentationPart;
                SlidePart slidePart = presentationPart.GetPartsOfType<SlidePart>().ElementAt(slideIndex);

                // Find the old image part
                ImagePart oldImagePart = null;
                string oldRelId = null;

                foreach (var blip in slidePart.Slide.Descendants<A.Blip>())
                {
                    oldRelId = blip.Embed;
                    oldImagePart = (ImagePart)slidePart.GetPartById(oldRelId);
                    if (oldImagePart != null)
                    {
                        break;
                    }
                }

                if (oldImagePart != null)
                {
                    // Add the new image part
                    ImagePart newImagePart = slidePart.AddImagePart(ImagePartType.Png);
                    using (MemoryStream stream = new MemoryStream(imageBytes))
                    {
                        newImagePart.FeedData(stream);
                    }

                    string newRelId = slidePart.GetIdOfPart(newImagePart);

                    // Replace the old image with the new image
                    foreach (var blip in slidePart.Slide.Descendants<A.Blip>())
                    {
                        if (blip.Embed == oldRelId)
                        {
                            blip.Embed = newRelId;
                        }
                    }

                    // Optionally delete the old image part to clean up
                    slidePart.DeletePart(oldImagePart);
                }
            }
        }

    }
}


