using System.ComponentModel.DataAnnotations;

namespace Automated_Certificate_Creation.Models
{
    public class CertificateGenerationViewModel
    {
        /// <summary>
        /// Certificate template file uploaded by user (.docx, .doc, .dotx, .rtf)
        /// Contains mail merge fields for certificate
        /// </summary>
        [Display(Name = "Certificate Template File")]
        public IFormFile? TemplateFile { get; set; }

        /// <summary>
        /// SQLite database file uploaded by user (.db, .sqlite, .sqlite3)
        /// Contains student records and related tables
        /// </summary>
        [Display(Name = "Student Database File")]
        public IFormFile? DatabaseFile { get; set; }

        /// <summary>
        /// Database table relationships for mail merge
        /// Format: "Students | string.Empty\nCourses | StudentID = %Students.StudentID%"
        /// Line-separated relationship commands
        /// </summary>
        [Display(Name = "Database Relationships")]
        public string? Relationships { get; set; }

        /// <summary>
        /// Type of certificate generation
        /// Options: "single" - One certificate with all records, "batch" - Individual certificates
        /// </summary>
        [Display(Name = "Certificate Type")]
        [Required]
        public string CertificateType { get; set; } = "single";

        /// <summary>
        /// Enable or disable digital signature feature
        /// </summary>
        [Display(Name = "Enable Digital Signature")]
        public bool EnableDigitalSign { get; set; } = false;

        /// <summary>
        /// Optional custom signature image file
        /// Supported formats: .png, .jpg, .jpeg, .gif, .bmp
        /// </summary>
        [Display(Name = "Signature Image File")]
        public IFormFile? SignatureImage { get; set; }

        /// <summary>
        /// Comma-separated keywords to identify signature placement locations
        /// Default: "Principal, Director, AuthorizedBy, Signature"
        /// </summary>
        [Display(Name = "Signature Keywords")]
        public string? SignatureKeywords { get; set; } = "Principal, Director, AuthorizedBy, Signature";
    }
}
