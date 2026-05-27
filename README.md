# Syncfusion ASP.NET Core – Automated Certificate Generation Demo

This repository contains a complete showcase sample demonstrating how to build an **Automated Certificate Generation System** using **Syncfusion DocIO**, **Syncfusion DocIORenderer**, **Syncfusion PDF**, and **Syncfusion Smart Data Extractor** libraries in an ASP.NET Core MVC application. The sample illustrates how educational institutions can streamline certificate creation by merging SQLite database records with Word templates, supporting both single and batch document generation with digital signature capabilities.

---

## 📁 Project Structure

```
├── Controllers/
│   └── HomeController.cs
├── Models/
│   └── CertificateGenerationViewModel.cs
├── Views/
│   ├── Home/
│   │   └── Index.cshtml
│   └── Shared/
├── wwwroot/
│   └── Data/
│       ├── Template.docx
│       ├── CertificateDatabase.db
├── PDF.pfx
├── Signature.png
└── README.md
```

---

## 🚀 Getting Started

### Prerequisites

- [.NET 7.0 SDK](https://dotnet.microsoft.com/download) or later
- Visual Studio 2022 or VS Code
- A valid **Syncfusion License Key** (or use the free Community License)

### 1. Clone the Repository

```bash
git clone https://github.com/SyncfusionExamples/DocIO-Automated-Certificate-Creation
cd Automated-Certificate-Creation
```

### 2. Install Dependencies

Restore all NuGet packages:

```bash
dotnet restore
```

### 3. Add Syncfusion License Key

In your `Program.cs`, register your Syncfusion license:

```csharp
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("YOUR_LICENSE_KEY");
```

### 4. Run the Application

```bash
dotnet run
```

Open your browser and navigate to:
```
https://localhost:5001
```

---

## 📋 How to Use

### Basic Workflow

1. **Upload Certificate Template**
   - Upload your Word template with merge fields
   - Drag & drop supported
   - Supports .DOCX, .DOC, .RTF formats
   - Or use default template provided

2. **Upload Student Database**
   - Upload SQLite database file containing student records
   - Supports .DB, .SQLITE, .SQLITE3 formats
   - Database should contain tables with student information
   - Or use default database file

3. **Configure Database Relationships**
   - Define table relationships using pipe-separated format
   - One relationship command per line
   - Format: `TableName | Relationship`
   - Example:
     ```
     Students | string.Empty
     Courses | StudentID = %Students.StudentID%
     Instructors | CourseID = %Courses.CourseID%
     Grades | StudentID = %Students.StudentID%
     ```

4. **Select Certificate Type**
   - **Single Certificate** - Generate one certificate with all records combined
   - **Batch Certificates** - Generate individual certificates for each student (ZIP archive)

5. **Configure Digital Signature (Optional)**
   - Enable authorized signature option
   - Upload signature image (PNG, JPG, etc.)
   - Enter placement keywords (e.g., "Principal, Director, Sign, Signature")
   - Signatures will automatically appear above matching keywords in the PDF

6. **Generate Certificates**
   - Click "Generate Certificate(s) as PDF" button
   - Download automatically starts
   - Single PDF file or ZIP archive based on selection

---

## 🎯 Use Cases

### Education Document Scenarios

- **Graduation Certificates** - Create personalized certificates from template and student database
- **Course Completion Certificates** - Generate certificates for completed courses
- **Achievement Awards** - Produce award certificates with student achievements
- **Bulk Certificate Generation** - Generate hundreds of certificates in a single operation
- **Transcript Generation** - Create academic transcripts with relational data from multiple tables
- **QR Code Integration** - Embed verification QR codes in certificates

---

## 🔗 Resources

- [Syncfusion DocIO Getting Started](https://help.syncfusion.com/document-processing/word/word-library/net/getting-started)
- [Mail Merge Documentation](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-mailmerge)
- [Word to PDF Conversion](https://help.syncfusion.com/document-processing/word/conversions/word-to-pdf/overview)
- [Smart Data Extractor](https://help.syncfusion.com/document-processing/data-extraction/smart-data-extractor/overview)
- [PDF Digital Signatures](https://help.syncfusion.com/document-processing/pdf/pdf-library/net/working-with-digitalsignature)
- [PDF Barcode Generation](https://help.syncfusion.com/document-processing/pdf/pdf-library/net/working-with-barcode)

---

## 📣 Try It Out

Clone the repository, run the sample, and discover how **Syncfusion DocIO** can revolutionize certificate generation in your educational institution.

### Customization

This sample application is provided as a reference implementation and can be freely customized to suit your specific business requirements.

You can modify the templates, database schema, merge logic, signature placement, and output formats based on your use case. If you have any questions, need clarification, or require assistance while customizing this sample, please feel free to contact our [Syncfusion Support Team](https://support.syncfusion.com/support/tickets/create) for guidance.

---


## 📄 License and Copyright

> This is a commercial product and requires a paid license for possession or use. Syncfusion® licensed software, including this component, is subject to the terms and conditions of Syncfusion®. To acquire a license, visit https://www.syncfusion.com/account/downloads.

Are you already a Syncfusion user? You can download the product setup [here](https://www.syncfusion.com/account/downloads). If you're not yet a Syncfusion user, you can download a [30-day free trial](https://www.syncfusion.com/downloads).

---

## 📞 Support

For technical support and questions:
- [Syncfusion Support Portal](https://support.syncfusion.com/support/tickets/create)
- [Documentation](https://help.syncfusion.com/)
- [Community Forums](https://www.syncfusion.com/forums)

---