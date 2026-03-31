namespace PDnDocuments;

public static class AppConstants
{
    public const string DefaultStatus = "Черновик";

    public static readonly string[] DocumentStatuses =
    [
        "Черновик",
        "На согласовании",
        "Действует",
        "На пересмотре",
        "Архив",
    ];

    public static readonly string[] TransferOptions =
    [
        "Не указано",
        "Нет",
        "Да",
    ];
}

public sealed class AppState
{
    public string Version { get; set; } = "1.0";

    public SequenceState Sequence { get; set; } = new();

    public List<CompanyRecord> Companies { get; set; } = [];

    public List<SectionRecord> Sections { get; set; } = [];

    public List<DocumentRecord> Documents { get; set; } = [];

    public List<SopdRecord> SopdRecords { get; set; } = [];

    public List<DocumentHistoryEntry> DocumentHistory { get; set; } = [];
}

public sealed class SequenceState
{
    public int NextCompanyId { get; set; } = 1;

    public int NextSectionId { get; set; } = 1;

    public int NextDocumentId { get; set; } = 1;

    public int NextSopdId { get; set; } = 1;

    public int NextHistoryId { get; set; } = 1;
}

public sealed class CompanyRecord
{
    public int Id { get; set; }

    public string Name { get; set; } = string.Empty;

    public string? SopdFilePath { get; set; }

    public string CreatedAt { get; set; } = string.Empty;

    public string CreatedBy { get; set; } = string.Empty;
}

public sealed class SectionRecord
{
    public int Id { get; set; }

    public int CompanyId { get; set; }

    public string Name { get; set; } = string.Empty;

    public int SortOrder { get; set; }

    public string CreatedAt { get; set; } = string.Empty;

    public string CreatedBy { get; set; } = string.Empty;
}

public sealed class DocumentRecord
{
    public int Id { get; set; }

    public int CompanyId { get; set; }

    public string Title { get; set; } = string.Empty;

    public string Status { get; set; } = AppConstants.DefaultStatus;

    public string? PdfPath { get; set; }

    public string? OfficePath { get; set; }

    public string Comment { get; set; } = string.Empty;

    public bool NeedsOffice { get; set; }

    public string? ReviewDue { get; set; }

    public string? AcceptDate { get; set; }

    public int SortOrder { get; set; }

    public List<int> SectionIds { get; set; } = [];

    public string CreatedAt { get; set; } = string.Empty;

    public string CreatedBy { get; set; } = string.Empty;

    public string UpdatedAt { get; set; } = string.Empty;

    public string UpdatedBy { get; set; } = string.Empty;
}

public sealed class SopdRecord
{
    public int Id { get; set; }

    public int CompanyId { get; set; }

    public string ConsentType { get; set; } = string.Empty;

    public string Purpose { get; set; } = string.Empty;

    public string LegalBasis { get; set; } = string.Empty;

    public string PDCategories { get; set; } = string.Empty;

    public string DataSubjects { get; set; } = string.Empty;

    public string PDList { get; set; } = string.Empty;

    public string ProcessingOperations { get; set; } = string.Empty;

    public string ProcessingMethod { get; set; } = string.Empty;

    public string ThirdPartyTransfer { get; set; } = "Не указано";

    public string TransferTo { get; set; } = string.Empty;

    public string Description { get; set; } = string.Empty;

    public string? AttachmentPath { get; set; }

    public string ValidityPeriod { get; set; } = string.Empty;

    public int SortOrder { get; set; }

    public string CreatedAt { get; set; } = string.Empty;

    public string CreatedBy { get; set; } = string.Empty;

    public string UpdatedAt { get; set; } = string.Empty;

    public string UpdatedBy { get; set; } = string.Empty;
}

public sealed class DocumentHistoryEntry
{
    public int Id { get; set; }

    public int DocumentId { get; set; }

    public string EventType { get; set; } = string.Empty;

    public string EventText { get; set; } = string.Empty;

    public string CreatedAt { get; set; } = string.Empty;

    public string CreatedBy { get; set; } = string.Empty;
}
