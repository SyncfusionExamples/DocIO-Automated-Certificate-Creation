using Automated_Certificate_Creation.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.Sqlite;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Drawing;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Pdf.Security;
using Syncfusion.SmartDataExtractor;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.IO.Compression;

namespace Automated_Certificate_Creation.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _hostingEnvironment;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment hostingEnvironment)
        {
            _logger = logger;
            _hostingEnvironment = hostingEnvironment;
        }

        public IActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Main action to generate property agreement from template and database
        /// </summary>
        public IActionResult GenerateCertificate(CertificateGenerationViewModel model)
        {
            try
            {
                // Step 1: Load Word template (uploaded or default)
                Stream wordStream = GetWordDocument(model.TemplateFile);
                if (wordStream == null)
                {
                    ViewBag.Message = "Failed to load Word template.";
                    return View("Index");
                }

                // Step 2: Load SQLite Database (uploaded or default)
                string databasePath = GetDatabaseFile(model.DatabaseFile);
                if (string.IsNullOrEmpty(databasePath))
                {
                    ViewBag.Message = "Failed to load SQLite database.";
                    return View("Index");
                }

                // Step 3: Create DataSet from SQLite database
                DataSet dataSet = CreateDataSetFromSQLite(databasePath);
                if (dataSet == null || dataSet.Tables.Count == 0)
                {
                    ViewBag.Message = "No tables found in the database.";
                    return View("Index");
                }

                // Step 4: Parse relationship commands from user input
                ArrayList commands = ParseRelationships(model.Relationships);

                // Step 5: Perform mail merge
                using (WordDocument document = new WordDocument(wordStream, FormatType.Automatic))
                {
                    ExecuteMailMerge(document, dataSet, commands);

                    // Step 6: Generate output based on certificate type
                    return GenerateOutput(document, model);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating certificate");
                ViewBag.Message = $"Error: {ex.Message}";
                return View("Index");
            }
        }

        /// <summary>
        /// Creates a DataSet from SQLite database by dynamically discovering all tables
        /// </summary>
        private DataSet CreateDataSetFromSQLite(string dbFilePath)
        {
            DataSet dataSet = new DataSet();
            string connectionString = $"Data Source={dbFilePath};";

            try
            {
                using (SqliteConnection conn = new SqliteConnection(connectionString))
                {
                    conn.Open();
                    _logger.LogInformation("SQLite database connection opened successfully.");

                    // Query to get all user tables (excluding SQLite system tables)
                    string tableQuery = "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name;";

                    List<string> tableNames = new List<string>();

                    using (SqliteCommand cmd = new SqliteCommand(tableQuery, conn))
                    using (SqliteDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            tableNames.Add(reader.GetString(0));
                        }
                    }

                    _logger.LogInformation($"Found {tableNames.Count} tables in database.");

                    // Load each table into DataSet using SqliteDataReader
                    foreach (string tableName in tableNames)
                    {
                        try
                        {
                            string selectQuery = $"SELECT * FROM [{tableName}]";

                            using (SqliteCommand selectCmd = new SqliteCommand(selectQuery, conn))
                            using (SqliteDataReader reader = selectCmd.ExecuteReader())
                            {
                                DataTable table = new DataTable(tableName);
                                table.Load(reader);  // Load data from reader into DataTable
                                dataSet.Tables.Add(table);
                                _logger.LogInformation($"Loaded table '{tableName}' with {table.Rows.Count} rows.");
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, $"Error loading table '{tableName}'");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error loading SQLite database: {dbFilePath}");
                throw new Exception($"Database connection error: {ex.Message}", ex);
            }

            return dataSet;
        }
        /// <summary>
        /// Executes mail merge based on Word template structure analysis
        /// </summary>
        private void ExecuteMailMerge(WordDocument document, DataSet dataSet, ArrayList commands)
        {
            if (dataSet.Tables.Count == 0)
            {
                _logger.LogWarning("No tables in DataSet");
                return;
            }

            try
            {
                document.MailMerge.StartAtNewPage = true;

                if (dataSet.Tables.Count == 1)
                {
                    DataTable table = dataSet.Tables[0];

                    if (table.Rows.Count > 1)
                    {
                        _logger.LogInformation($"ExecuteGroup: {table.Rows.Count} records");
                        document.MailMerge.ExecuteGroup(table);
                    }
                    else
                    {
                        _logger.LogInformation("Execute: Single record");
                        document.MailMerge.StartAtNewPage = false;
                        document.MailMerge.Execute(table);
                    }
                }
                else
                {
                    bool hasNestedStructure = HasNestedStructure(document, dataSet);

                    if (hasNestedStructure && commands != null)
                    {
                        _logger.LogInformation("ExecuteNestedGroup: Nested structure detected");
                        document.MailMerge.ExecuteNestedGroup(dataSet, commands);
                    }
                    else
                    {
                        _logger.LogInformation("ExecuteNestedGroup: Flat structure");
                        commands = new ArrayList();
                        foreach (DataTable table in dataSet.Tables)
                        {
                            commands.Add(new DictionaryEntry(table.TableName, string.Empty));
                        }
                        document.MailMerge.ExecuteNestedGroup(dataSet, commands);
                    }
                }

                document.UpdateDocumentFields();
                _logger.LogInformation("Mail merge completed successfully");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Mail merge failed");
                throw new Exception($"Mail merge failed: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Checks if Word template has nested merge regions
        /// </summary>
        private bool HasNestedStructure(WordDocument document, DataSet dataSet)
        {
            try
            {
                string[] groupNames = document.MailMerge.GetMergeGroupNames();

                if (groupNames == null || groupNames.Length < 2)
                {
                    _logger.LogInformation("No nested structure: Less than 2 groups in template");
                    return false;
                }

                string rootGroup = groupNames[0];

                DataTable parentTable = dataSet.Tables
                    .Cast<DataTable>()
                    .FirstOrDefault(t => string.Equals(t.TableName, rootGroup, StringComparison.OrdinalIgnoreCase));

                if (parentTable == null)
                {
                    _logger.LogWarning($"Root group '{rootGroup}' not found in DataSet");
                    return false;
                }

                foreach (string childGroup in groupNames.Skip(1))
                {
                    DataTable childTable = dataSet.Tables
                        .Cast<DataTable>()
                        .FirstOrDefault(t => string.Equals(t.TableName, childGroup, StringComparison.OrdinalIgnoreCase));

                    if (childTable == null)
                        continue;

                    string commonColumn = FindCommonColumn(parentTable, childTable);

                    if (!string.IsNullOrEmpty(commonColumn))
                    {
                        _logger.LogInformation($"Nested structure found: {parentTable.TableName}.{commonColumn} -> {childTable.TableName}.{commonColumn}");
                        return true;
                    }
                }

                _logger.LogInformation("No nested structure: No common columns found between tables");
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking nested structure");
                return false;
            }
        }

        /// <summary>
        /// Finds common column name between two tables
        /// </summary>
        private string FindCommonColumn(DataTable parentTable, DataTable childTable)
        {
            var parentColumns = parentTable.Columns.Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var childColumns = childTable.Columns.Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            string commonColumn = parentColumns.Intersect(childColumns, StringComparer.OrdinalIgnoreCase).FirstOrDefault();

            return commonColumn;
        }

        /// <summary>
        /// Generates output document based on certificate type (single or batch)
        /// </summary>
        private IActionResult GenerateOutput(WordDocument document, CertificateGenerationViewModel model)
        {
            bool isSingleCertificate = model.CertificateType == "single";

            if (isSingleCertificate)
            {
                // Generate single PDF with optional signature
                MemoryStream pdfStream = SaveAsPDF(document);

                if (model.EnableDigitalSign)
                {
                    pdfStream = ApplyDigitalSignatureIfEnabled(pdfStream, model.SignatureImage, model.SignatureKeywords, model.EnableDigitalSign);
                }

                return File(pdfStream, "application/pdf", "Certificate.pdf");
            }
            else
            {
                // Generate batch certificates as ZIP
                byte[] zipBytes = SplitByPageBreak(document, model.SignatureImage, model.SignatureKeywords, model.EnableDigitalSign);

                if (zipBytes != null && zipBytes.Length > 0)
                {
                    return File(zipBytes, "application/zip", "Certificates.zip");
                }
                else
                {
                    // Fallback to single PDF
                    MemoryStream pdfStream = SaveAsPDF(document);

                    if (model.EnableDigitalSign)
                    {
                        pdfStream = ApplyDigitalSignatureIfEnabled(pdfStream, model.SignatureImage, model.SignatureKeywords, model.EnableDigitalSign);
                    }

                    return File(pdfStream, "application/pdf", "Certificate.pdf");
                }
            }
        }
        /// <summary>
        /// Splits document by page breaks using bookmarks and returns ZIP bytes
        /// Each section between page breaks becomes a separate PDF
        /// </summary>
        private byte[] SplitByPageBreak(WordDocument wordDocument, IFormFile signatureImage, string signatureKeywords, bool enableDigitalSign)
        {
            // Find all page breaks in the document
            List<Entity> entities = wordDocument.FindAllItemsByProperty(EntityType.Break, "BreakType", "PageBreak");
            if (entities == null || entities.Count == 0)
                return null;

            WSection section = wordDocument.Sections[0];
            WTextBody body = section.Body;
            int bookmarkIndex = 1;
            // Step 1: Insert a NEW paragraph at the very beginning with BookmarkStart
            WParagraph firstBookmarkPara = new WParagraph(wordDocument);
            firstBookmarkPara.AppendBookmarkStart($"Page_Bookmark_{bookmarkIndex}");
            body.ChildEntities.Insert(0, firstBookmarkPara);

            // Step 2: Iterate page break entities → insert bookmark paragraph directly after each
            foreach (Entity entity in entities)
            {
                WParagraph breakParagraph = entity.Owner as WParagraph;

                if (breakParagraph == null) continue;

                // Get the current index of this paragraph in the body
                int paraIndex = body.ChildEntities.IndexOf(breakParagraph);

                if (paraIndex < 0) continue;

                // Insert new paragraph right after the page break paragraph
                // Close current bookmark and open next bookmark in same paragraph
                WParagraph bookmarkPara = new WParagraph(wordDocument);
                bookmarkPara.AppendBookmarkEnd($"Page_Bookmark_{bookmarkIndex}");
                bookmarkIndex++;
                bookmarkPara.AppendBookmarkStart($"Page_Bookmark_{bookmarkIndex}");
                body.ChildEntities.Insert(paraIndex + 1, bookmarkPara);
            }

            // Step 3: Insert a NEW paragraph at the very end with BookmarkEnd
            WParagraph lastBookmarkPara = new WParagraph(wordDocument);
            lastBookmarkPara.AppendBookmarkEnd($"Page_Bookmark_{bookmarkIndex}");
            body.ChildEntities.Add(lastBookmarkPara);
            // Step 4: Create ZIP file and convert each bookmarked section to PDF
            var zipStream = new MemoryStream();
            using (var zip = new ZipArchive(zipStream, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
            {
                for (int i = 1; i <= bookmarkIndex; i++)
                {
                    try
                    {
                        // Navigate to each bookmark section
                        BookmarksNavigator navigator = new BookmarksNavigator(wordDocument);
                        navigator.MoveToBookmark($"Page_Bookmark_{i}", true, true);
                        WordDocumentPart documentPart = navigator.GetContent();

                        if (documentPart == null) continue;

                        using (WordDocument extractedDoc = documentPart.GetAsWordDocument())
                        {

                            // Convert to PDF
                            using (DocIORenderer render = new DocIORenderer())
                            using (PdfDocument pdfDocument = render.ConvertToPDF(extractedDoc))
                            using (MemoryStream pdfStream = new MemoryStream())
                            {
                                pdfDocument.Save(pdfStream);
                                pdfStream.Position = 0;

                                // Apply digital signatures to each PDF
                                MemoryStream signedPdfStream = ApplyDigitalSignatureIfEnabled(pdfStream, signatureImage, signatureKeywords, enableDigitalSign);

                                var entry = zip.CreateEntry($"Document_{i}.pdf", CompressionLevel.Fastest);
                                using (var entryStream = entry.Open())
                                {
                                    signedPdfStream.CopyTo(entryStream);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log and continue to next section if one fails
                        _logger.LogError(ex, $"Error converting bookmark {i} to PDF");
                    }
                }
            }
            return zipStream.ToArray();
        }

        /// <summary>
        /// Retrieves Word template stream from uploaded file or default template
        /// </summary>
        private Stream GetWordDocument(IFormFile file)
        {
            // Case 1: User uploaded a template file
            if (file != null && file.Length > 0)
            {
                string extension = Path.GetExtension(file.FileName).ToLower();
                string[] supportedExtensions = { ".doc", ".docx", ".dot", ".dotx", ".dotm", ".docm", ".rtf" };

                if (supportedExtensions.Contains(extension))
                {
                    MemoryStream stream = new MemoryStream();
                    file.CopyTo(stream);
                    stream.Position = 0;
                    _logger.LogInformation("Using user-uploaded Word template.");
                    return stream;
                }
                else
                {
                    ViewBag.Message = "Please upload a valid Word document format.";
                    return null;
                }
            }
            else
            {
                // Case 2: Use default template
                string defaultFilePath = Path.Combine(_hostingEnvironment.WebRootPath, "Data", "PropertyAgreementTemplate.docx");

                if (System.IO.File.Exists(defaultFilePath))
                {
                    using (var fileStream = new FileStream(defaultFilePath, FileMode.Open, FileAccess.Read))
                    {
                        var memoryStream = new MemoryStream();
                        fileStream.CopyTo(memoryStream);
                        memoryStream.Position = 0;
                        _logger.LogInformation("Using default Word template.");
                        return memoryStream;
                    }
                }
                else
                {
                    _logger.LogError($"Default template not found at: {defaultFilePath}");
                    ViewBag.Message = "Default template file not found.";
                    return null;
                }
            }
        }

        /// <summary>
        /// Retrieves database file path from uploaded file or default database
        /// </summary>
        private string GetDatabaseFile(IFormFile databaseFile)
        {
            if (databaseFile != null && databaseFile.Length > 0)
            {
                string extension = Path.GetExtension(databaseFile.FileName).ToLower();

                if (extension == ".db" || extension == ".sqlite" || extension == ".sqlite3")
                {
                    string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + extension);

                    using (var fileStream = new FileStream(tempPath, FileMode.Create))
                    {
                        databaseFile.CopyTo(fileStream);
                    }

                    _logger.LogInformation($"Using user-uploaded database: {databaseFile.FileName}");
                    return tempPath;
                }
            }

            string defaultFilePath = Path.Combine(_hostingEnvironment.WebRootPath, "Data", "StudentDatabase.db");

            if (System.IO.File.Exists(defaultFilePath))
            {
                _logger.LogInformation("Using default database.");
                return defaultFilePath;
            }

            _logger.LogError($"Default database not found at: {defaultFilePath}");
            return null;
        }

        /// <summary>
        /// Creates a DataSet from an MDB/ACCDB file by dynamically discovering and loading all tables
        /// </summary>
        private DataSet CreateDataSetFromMDB(string mdbFilePath)
        {
            DataSet dataSet = new DataSet();

            

            return dataSet;
        }

        /// <summary>
        /// Parses user-provided relationship string into ArrayList of DictionaryEntry commands
        /// Format: Each line contains "TableName | Relationship"
        /// Example:
        /// Employees | string.Empty
        /// Customers | EmployeeID = %Employees.EmployeeID%
        /// Orders | CustomerID = %Customers.CustomerID%
        /// </summary>
        private ArrayList ParseRelationships(string relationships)
        {
            ArrayList commands = new ArrayList();

            if (string.IsNullOrWhiteSpace(relationships))
            {
                _logger.LogWarning("No relationships provided.");
                return null;
            }

            try
            {
                // Split by newline
                string[] lines = relationships
                    .Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(line => line.Trim())
                    .Where(line => !string.IsNullOrEmpty(line))
                    .ToArray();

                if (lines.Length == 0)
                {
                    _logger.LogWarning("No valid relationships found.");
                    return null;
                }

                // Process each line: TableName | Relationship
                foreach (string line in lines)
                {
                    if (!line.Contains("|"))
                    {
                        _logger.LogWarning($"Invalid line format (missing '|'): {line}");
                        continue;
                    }

                    string[] parts = line.Split('|');
                    if (parts.Length != 2)
                    {
                        _logger.LogWarning($"Invalid line format: {line}");
                        continue;
                    }

                    string tableName = parts[0].Trim();
                    string relationship = parts[1].Trim();

                    // Handle "string.Empty" or empty for parent table
                    if (relationship.Equals("string.Empty", StringComparison.OrdinalIgnoreCase) ||
                        string.IsNullOrEmpty(relationship))
                    {
                        relationship = string.Empty;
                    }

                    DictionaryEntry entry = new DictionaryEntry(tableName, relationship);
                    commands.Add(entry);

                    _logger.LogInformation($"Added: Table='{tableName}', Relation='{relationship}'");
                }

                if (commands.Count == 0)
                {
                    _logger.LogWarning("No valid relationships found.");
                    return null;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error parsing relationships");
                return null;
            }

            return commands;
        }


        /// <summary>
        /// Extracts table name from "TableName.ColumnName"
        /// </summary>
        private string ExtractTableName(string tableDotColumn)
        {
            if (tableDotColumn.Contains("."))
            {
                return tableDotColumn.Split('.')[0].Trim();
            }
            return tableDotColumn.Trim();
        }

        /// <summary>
        /// Extracts column name from "TableName.ColumnName"
        /// </summary>
        private string ExtractColumnName(string tableDotColumn)
        {
            if (tableDotColumn.Contains("."))
            {
                return tableDotColumn.Split('.')[1].Trim();
            }
            return tableDotColumn.Trim();
        }

        /// <summary>
        /// Converts Word document to PDF and returns PDF stream
        /// </summary>
        private MemoryStream SaveAsPDF(WordDocument wordDocument)
        {
            using (DocIORenderer renderer = new DocIORenderer())
            {
                using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
                {
                    MemoryStream pdfStream = new MemoryStream();
                    pdfDocument.Save(pdfStream);
                    pdfStream.Position = 0;
                    _logger.LogInformation("Word document converted to PDF successfully.");
                    return pdfStream;
                }
            }
        }

        /// <summary>
        /// Conditionally applies digital signatures to PDF based on enableDigitalSign flag
        /// </summary>
        private MemoryStream ApplyDigitalSignatureIfEnabled(MemoryStream inputStream, IFormFile signatureImage, string signatureKeywordsInput, bool enableDigitalSign)
        {
            Stream signatureStream = null;
            try
            {
                // Early exit if digital signature is not enabled
                if (!enableDigitalSign)
                {
                    _logger.LogInformation("Digital signature is not enabled. Returning PDF stream directly.");
                    inputStream.Position = 0;
                    return inputStream;
                }

                // Check if signature image is available
                signatureStream = GetSignatureImageStream(signatureImage);
                if (signatureStream == null)
                {
                    _logger.LogWarning("Digital signature enabled but no signature image found. Returning PDF stream directly.");
                    inputStream.Position = 0;
                    return inputStream;
                }

                // Apply digital signatures
                _logger.LogInformation("Applying digital signatures to PDF document.");

                // Initialize the extractor with required detection settings
                var extractor = new DataExtractor { EnableFormDetection = false, EnableTableDetection = true, ConfidenceThreshold = 0.6 };

                // Extract PDF document from the input stream
                inputStream.Position = 0;
                PdfLoadedDocument pdfDocument = extractor.ExtractDataAsPdfDocument(inputStream);

                // Add signatures
                AddSignaturesToPDFDocument(pdfDocument, signatureStream, signatureKeywordsInput);

                // Save PDF with signatures
                var outputMs = new MemoryStream();
                pdfDocument.Save(outputMs);
                pdfDocument.Close(true);
                outputMs.Position = 0;

                _logger.LogInformation("Digital signatures applied successfully.");
                return outputMs;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing PDF document with digital signatures");
                throw;
            }
            finally
            {
                signatureStream?.Dispose();
            }
        }

        /// <summary>
        /// Adds digital signatures to PDF document at locations matching specified keywords
        /// </summary>
        private void AddSignaturesToPDFDocument(PdfLoadedDocument pdfDocument, Stream signatureImageStream, string keywords)
        {
            // Use default keywords if none are provided
            string[] signatureKeywords = string.IsNullOrWhiteSpace(keywords)
                ? new[] { "Sign", "WITNESS", "AuthorizedSign", "Signature" }
                : keywords.Split(',').Select(k => k.Trim()).ToArray();

            _logger.LogInformation($"Applying digital signatures for keywords: {string.Join(", ", signatureKeywords)}");

            // Iterate through each page in the document
            for (int pageIndex = 0; pageIndex < pdfDocument.Pages.Count; pageIndex++)
            {
                PdfPageBase page = pdfDocument.Pages[pageIndex];
                TextLineCollection textLines;

                // Extract text lines from the page
                page.ExtractText(out textLines);

                // Iterate through each word on the page
                foreach (TextLine line in textLines.TextLine)
                {
                    foreach (TextWord word in line.WordCollection)
                    {
                        // Skip words that do not match any signature keyword
                        if (!signatureKeywords.Any(k => word.Text.Contains(k, StringComparison.Ordinal)))
                            continue;

                        // Calculate signature position above the keyword
                        RectangleF bounds = word.Bounds;
                        float signatureX = bounds.X;
                        float signatureY = bounds.Y - bounds.Height - 10;
                        float signatureWidth = 80;
                        float signatureHeight = 20;

                        try
                        {
                            // Load digital certificate
                            string certPath = Path.Combine(_hostingEnvironment.ContentRootPath, "PDF.pfx");

                            if (!System.IO.File.Exists(certPath))
                            {
                                _logger.LogWarning($"Certificate file not found at: {certPath}");
                                continue;
                            }

                            using System.IO.FileStream cert = new System.IO.FileStream(certPath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                            PdfCertificate pdfCert = new PdfCertificate(cert, "syncfusion");

                            // Create and configure the PDF signature
                            PdfSignature signature = new PdfSignature(pdfDocument, page, pdfCert, "Signature");
                            signature.Bounds = new RectangleF(signatureX, signatureY, signatureWidth, signatureHeight);

                            // Load signature image directly from stream
                            signatureImageStream.Position = 0;
                            PdfBitmap signatureImageBitmap = new PdfBitmap(signatureImageStream);
                            signature.Appearance.Normal.Graphics.DrawImage(signatureImageBitmap, 0, 0, signatureWidth, signatureHeight);

                            _logger.LogDebug($"Signature added at page {pageIndex + 1}, position ({signatureX}, {signatureY})");
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, $"Error adding signature at page {pageIndex + 1}");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets signature image as a stream from user upload or default signature
        /// </summary>
        private Stream GetSignatureImageStream(IFormFile signatureImage)
        {
            // If user provided an image, return its stream
            if (signatureImage != null && signatureImage.Length > 0)
            {
                _logger.LogInformation("Using user-provided signature image stream.");
                var memoryStream = new MemoryStream();
                signatureImage.OpenReadStream().CopyTo(memoryStream);
                memoryStream.Position = 0;
                return memoryStream;
            }

            // No user image - use default signature from project
            string defaultImagePath = Path.Combine(_hostingEnvironment.ContentRootPath, "Signature.png");

            if (System.IO.File.Exists(defaultImagePath))
            {
                _logger.LogInformation($"Using default signature image from: {defaultImagePath}");
                try
                {
                    var fileStream = new FileStream(defaultImagePath, FileMode.Open, FileAccess.Read);
                    var memoryStream = new MemoryStream();
                    fileStream.CopyTo(memoryStream);
                    fileStream.Dispose();
                    memoryStream.Position = 0;
                    return memoryStream;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"Error loading default signature image: {defaultImagePath}");
                    return null;
                }
            }

            _logger.LogWarning($"Default signature image not found at: {defaultImagePath}");
            return null;
        }


        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
