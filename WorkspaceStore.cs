using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace PDnDocuments;

public sealed class WorkspaceStore
{
    private const string SettingsFileName = "pddoc-csharp.settings.json";
    private const string LegacyManagedStorageFolderName = ".pddoc-data";
    private const string LegacyWorkspaceFolderName = "workspace";
    private const string ContainerFileName = ".pddoc-store";
    private const string RegistryEntryPath = "data/registry.json";

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true,
    };

    private readonly string _applicationRootPath;
    private readonly string _currentUser;
    private readonly string _settingsPath;
    public WorkspaceStore(string applicationRootPath, string currentUser)
    {
        _applicationRootPath = Path.GetFullPath(applicationRootPath);
        _currentUser = currentUser;
        _settingsPath = Path.Combine(_applicationRootPath, SettingsFileName);
        CurrentWorkspacePath = ResolveStartupWorkspace();
        EnsureContainerReady();
        EnsureSettingsFile();
    }

    public string CurrentWorkspacePath { get; private set; }

    public string ContainerFilePath => Path.Combine(CurrentWorkspacePath, ContainerFileName);

    public string DataDirectoryPath => TempCacheRootPath;

    public string StorageDirectoryPath => ContainerFilePath;

    public string RegistryFilePath => $"{ContainerFilePath}::{RegistryEntryPath}";

    public AppState LoadState()
    {
        EnsureContainerReady();

        try
        {
            var json = ReadTextEntry(RegistryEntryPath);
            if (string.IsNullOrWhiteSpace(json))
            {
                var emptyState = CreateEmptyState();
                SaveState(emptyState);
                return emptyState;
            }

            var state = JsonSerializer.Deserialize<AppState>(json, JsonOptions) ?? CreateEmptyState();
            NormalizeState(state);
            return state;
        }
        catch
        {
            var recoveredState = CreateEmptyState();
            SaveState(recoveredState);
            return recoveredState;
        }
    }

    public void SaveState(AppState state)
    {
        EnsureContainerReady();
        NormalizeState(state);
        var json = JsonSerializer.Serialize(state, JsonOptions);
        WriteTextEntry(RegistryEntryPath, json);
    }

    public void ChangeWorkspace(string newWorkspacePath)
    {
        if (string.IsNullOrWhiteSpace(newWorkspacePath))
        {
            throw new ArgumentException("Рабочая папка не указана.", nameof(newWorkspacePath));
        }

        ClearTemporaryFiles();
        CurrentWorkspacePath = Path.GetFullPath(newWorkspacePath.Trim());
        EnsureContainerReady();
        SaveSettings();
    }

    public string CopyDocumentPdf(int companyId, int documentId, string sourcePath) =>
        CopyFileIntoContainer(sourcePath, "storage", "companies", CompanyFolder(companyId), "documents", DocumentFolder(documentId), "pdf");

    public string CopyDocumentOffice(int companyId, int documentId, string sourcePath) =>
        CopyFileIntoContainer(sourcePath, "storage", "companies", CompanyFolder(companyId), "documents", DocumentFolder(documentId), "office");

    public string CopyCompanySopdFile(int companyId, string sourcePath) =>
        CopyFileIntoContainer(sourcePath, "storage", "companies", CompanyFolder(companyId), "company-sopd");

    public string CopySopdAttachment(int companyId, int sopdId, string sourcePath) =>
        CopyFileIntoContainer(sourcePath, "storage", "companies", CompanyFolder(companyId), "sopd", SopdFolder(sopdId));

    public string? ResolveAbsolutePath(string? relativePath)
    {
        if (string.IsNullOrWhiteSpace(relativePath))
        {
            return null;
        }

        if (Path.IsPathRooted(relativePath))
        {
            return relativePath;
        }

        var entryName = NormalizeEntryPath(relativePath);
        if (!ArchiveEntryExists(entryName))
        {
            return null;
        }

        var extractedPath = BuildCachePath(entryName);
        Directory.CreateDirectory(Path.GetDirectoryName(extractedPath)!);

        using var archive = OpenArchive(ZipArchiveMode.Read);
        var entry = archive.GetEntry(entryName);
        if (entry is null)
        {
            return null;
        }

        using var input = entry.Open();
        using var output = new FileStream(extractedPath, FileMode.Create, FileAccess.Write, FileShare.Read);
        input.CopyTo(output);
        return extractedPath;
    }

    public bool RelativePathExists(string? relativePath)
    {
        if (string.IsNullOrWhiteSpace(relativePath))
        {
            return false;
        }

        return Path.IsPathRooted(relativePath)
            ? File.Exists(relativePath)
            : ArchiveEntryExists(relativePath);
    }

    public void DeleteRelativeFile(string? relativePath)
    {
        if (string.IsNullOrWhiteSpace(relativePath))
        {
            return;
        }

        if (Path.IsPathRooted(relativePath))
        {
            if (File.Exists(relativePath))
            {
                File.Delete(relativePath);
            }

            return;
        }

        EnsureContainerReady();
        var entryName = NormalizeEntryPath(relativePath);
        using var archive = OpenArchive(ZipArchiveMode.Update);
        archive.GetEntry(entryName)?.Delete();

        var cachedPath = BuildCachePath(entryName);
        if (File.Exists(cachedPath))
        {
            File.Delete(cachedPath);
        }
    }

    public void OpenWorkspaceInExplorer()
    {
        Directory.CreateDirectory(CurrentWorkspacePath);
        if (File.Exists(ContainerFilePath))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"/select,\"{ContainerFilePath}\"",
                UseShellExecute = true,
            });
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = CurrentWorkspacePath,
            UseShellExecute = true,
        });
    }

    public void CreateBackup(string destinationPath)
    {
        if (string.IsNullOrWhiteSpace(destinationPath))
        {
            throw new ArgumentException("Путь для резервной копии не указан.", nameof(destinationPath));
        }

        EnsureContainerReady();

        var fullPath = Path.GetFullPath(destinationPath.Trim());
        var directory = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrWhiteSpace(directory))
        {
            throw new InvalidOperationException("Не удалось определить папку для резервной копии.");
        }

        Directory.CreateDirectory(directory);
        File.Copy(ContainerFilePath, fullPath, overwrite: true);
    }

    public void RestoreBackup(string sourcePath)
    {
        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
        {
            throw new FileNotFoundException("Не удалось найти файл резервной копии.", sourcePath);
        }

        ValidateBackupFile(sourcePath);

        Directory.CreateDirectory(CurrentWorkspacePath);
        ClearTemporaryFiles();

        var tempRestorePath = Path.Combine(CurrentWorkspacePath, $"{ContainerFileName}.restore");
        if (File.Exists(tempRestorePath))
        {
            File.Delete(tempRestorePath);
        }

        File.Copy(sourcePath, tempRestorePath, overwrite: true);
        if (File.Exists(ContainerFilePath))
        {
            File.Delete(ContainerFilePath);
        }

        File.Move(tempRestorePath, ContainerFilePath);
        HidePathIfPossible(ContainerFilePath);
        CleanupLegacyStorageIfPossible();
    }

    public void ClearTemporaryFiles()
    {
        try
        {
            if (Directory.Exists(TempCacheRootPath))
            {
                Directory.Delete(TempCacheRootPath, recursive: true);
            }
        }
        catch
        {
            // Temp cleanup is best-effort only.
        }
    }

    private AppState CreateEmptyState() => new();

    private void EnsureContainerReady()
    {
        Directory.CreateDirectory(CurrentWorkspacePath);
        MigrateLegacyStorageIfNeeded();

        if (!File.Exists(ContainerFilePath))
        {
            using var archive = ZipFile.Open(ContainerFilePath, ZipArchiveMode.Create);
        }

        HidePathIfPossible(ContainerFilePath);
        CleanupLegacyStorageIfPossible();
    }

    private void MigrateLegacyStorageIfNeeded()
    {
        if (File.Exists(ContainerFilePath))
        {
            return;
        }

        var sourceDirectory = ResolveLegacySourceDirectory();
        if (sourceDirectory is null)
        {
            return;
        }

        using (var archive = ZipFile.Open(ContainerFilePath, ZipArchiveMode.Create))
        {
            foreach (var filePath in Directory.EnumerateFiles(sourceDirectory, "*", SearchOption.AllDirectories))
            {
                var relativePath = Path.GetRelativePath(sourceDirectory, filePath)
                    .Replace(Path.DirectorySeparatorChar, '/');
                archive.CreateEntryFromFile(filePath, NormalizeEntryPath(relativePath), CompressionLevel.Optimal);
            }
        }

        HidePathIfPossible(ContainerFilePath);
    }

    private string? ResolveLegacySourceDirectory()
    {
        var managedDirectory = Path.Combine(CurrentWorkspacePath, LegacyManagedStorageFolderName);
        if (Directory.Exists(managedDirectory))
        {
            return managedDirectory;
        }

        var legacyWorkspaceDirectory = Path.Combine(CurrentWorkspacePath, LegacyWorkspaceFolderName);
        if (Directory.Exists(legacyWorkspaceDirectory))
        {
            return legacyWorkspaceDirectory;
        }

        return null;
    }

    private string ResolveStartupWorkspace()
    {
        var defaultWorkspace = _applicationRootPath;

        try
        {
            if (!File.Exists(_settingsPath))
            {
                return defaultWorkspace;
            }

            var json = File.ReadAllText(_settingsPath, Encoding.UTF8);
            var settings = JsonSerializer.Deserialize<WorkspaceSettings>(json, JsonOptions);
            if (string.IsNullOrWhiteSpace(settings?.WorkspacePath))
            {
                return defaultWorkspace;
            }

            var configuredPath = Path.GetFullPath(settings.WorkspacePath);
            var trimmedConfiguredPath = configuredPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            var leaf = Path.GetFileName(trimmedConfiguredPath);
            var looksLikeLegacyStorage =
                Directory.Exists(configuredPath) &&
                File.Exists(Path.Combine(configuredPath, "data", "registry.json")) &&
                Directory.Exists(Path.Combine(configuredPath, "storage"));
            if (string.Equals(leaf, LegacyManagedStorageFolderName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(leaf, LegacyWorkspaceFolderName, StringComparison.OrdinalIgnoreCase) ||
                looksLikeLegacyStorage)
            {
                return Path.GetDirectoryName(trimmedConfiguredPath) ?? defaultWorkspace;
            }

            return configuredPath;
        }
        catch
        {
            return defaultWorkspace;
        }
    }

    private void SaveSettings()
    {
        var settings = new WorkspaceSettings
        {
            WorkspacePath = CurrentWorkspacePath,
        };

        var json = JsonSerializer.Serialize(settings, JsonOptions);
        File.WriteAllText(_settingsPath, json, Encoding.UTF8);
        HidePathIfPossible(_settingsPath);
    }

    private void EnsureSettingsFile()
    {
        if (!File.Exists(_settingsPath))
        {
            SaveSettings();
            return;
        }

        try
        {
            var json = File.ReadAllText(_settingsPath, Encoding.UTF8);
            var settings = JsonSerializer.Deserialize<WorkspaceSettings>(json, JsonOptions);
            var configuredPath = string.IsNullOrWhiteSpace(settings?.WorkspacePath)
                ? null
                : Path.GetFullPath(settings.WorkspacePath);
            if (string.Equals(configuredPath, CurrentWorkspacePath, StringComparison.OrdinalIgnoreCase))
            {
                HidePathIfPossible(_settingsPath);
                return;
            }
        }
        catch
        {
            // Rewrite broken settings below.
        }

        SaveSettings();
    }

    private void CleanupLegacyStorageIfPossible()
    {
        if (!File.Exists(ContainerFilePath))
        {
            return;
        }

        TryDeleteLegacyDirectory(Path.Combine(CurrentWorkspacePath, LegacyManagedStorageFolderName));
        TryDeleteLegacyDirectory(Path.Combine(CurrentWorkspacePath, LegacyWorkspaceFolderName));
    }

    private void TryDeleteLegacyDirectory(string path)
    {
        try
        {
            var fullPath = Path.GetFullPath(path);
            var expectedPath = Path.Combine(CurrentWorkspacePath, Path.GetFileName(path));
            if (!string.Equals(fullPath, expectedPath, StringComparison.OrdinalIgnoreCase) || !Directory.Exists(fullPath))
            {
                return;
            }

            Directory.Delete(fullPath, recursive: true);
        }
        catch
        {
            // Cleanup is best-effort only.
        }
    }

    private string CopyFileIntoContainer(string sourcePath, params string[] pathParts)
    {
        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
        {
            throw new FileNotFoundException("Не удалось найти исходный файл.", sourcePath);
        }

        EnsureContainerReady();

        var sourceFileName = Path.GetFileName(sourcePath);
        var stem = SanitizeFileName(Path.GetFileNameWithoutExtension(sourceFileName));
        var extension = Path.GetExtension(sourceFileName);
        var entryDirectory = NormalizeEntryPath(Path.Combine(pathParts));
        var targetName = $"{DateTime.Now:yyyyMMdd_HHmmss}_{stem}{extension}";
        var entryPath = $"{entryDirectory}/{targetName}";
        var counter = 1;

        using var archive = OpenArchive(ZipArchiveMode.Update);
        while (archive.GetEntry(entryPath) is not null)
        {
            targetName = $"{DateTime.Now:yyyyMMdd_HHmmss}_{stem}_{counter}{extension}";
            entryPath = $"{entryDirectory}/{targetName}";
            counter++;
        }

        var entry = archive.CreateEntry(entryPath, CompressionLevel.Optimal);
        using var input = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var output = entry.Open();
        input.CopyTo(output);
        return entryPath;
    }

    private bool ArchiveEntryExists(string relativePath)
    {
        if (!File.Exists(ContainerFilePath))
        {
            return false;
        }

        var entryName = NormalizeEntryPath(relativePath);
        using var archive = OpenArchive(ZipArchiveMode.Read);
        return archive.GetEntry(entryName) is not null;
    }

    private string? ReadTextEntry(string entryPath)
    {
        if (!File.Exists(ContainerFilePath))
        {
            return null;
        }

        using var archive = OpenArchive(ZipArchiveMode.Read);
        using var reader = OpenEntryReader(archive, entryPath);
        return reader?.ReadToEnd();
    }

    private void WriteTextEntry(string entryPath, string text)
    {
        EnsureContainerReady();
        using var archive = OpenArchive(ZipArchiveMode.Update);
        archive.GetEntry(NormalizeEntryPath(entryPath))?.Delete();
        var entry = archive.CreateEntry(NormalizeEntryPath(entryPath), CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), Encoding.UTF8);
        writer.Write(text);
    }

    private StreamReader? OpenEntryReader(ZipArchive archive, string entryPath)
    {
        var entry = archive.GetEntry(NormalizeEntryPath(entryPath));
        return entry is null ? null : new StreamReader(entry.Open(), Encoding.UTF8);
    }

    private static void ValidateBackupFile(string sourcePath)
    {
        try
        {
            using var archive = ZipFile.OpenRead(sourcePath);
            if (archive.GetEntry(RegistryEntryPath) is null)
            {
                throw new InvalidDataException("В резервной копии не найден файл реестра.");
            }
        }
        catch (InvalidDataException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new InvalidDataException("Файл резервной копии поврежден или имеет неверный формат.", ex);
        }
    }

    private ZipArchive OpenArchive(ZipArchiveMode mode)
    {
        var fileMode = mode == ZipArchiveMode.Read ? FileMode.Open : FileMode.Open;
        var fileAccess = mode == ZipArchiveMode.Read ? FileAccess.Read : FileAccess.ReadWrite;
        var fileShare = mode == ZipArchiveMode.Read ? FileShare.ReadWrite : FileShare.None;
        var stream = new FileStream(ContainerFilePath, fileMode, fileAccess, fileShare);
        return new ZipArchive(stream, mode, leaveOpen: false);
    }

    private string BuildCachePath(string entryName)
    {
        var normalizedPath = entryName.Replace('/', Path.DirectorySeparatorChar);
        var fullPath = Path.GetFullPath(Path.Combine(TempCacheRootPath, normalizedPath));
        if (!fullPath.StartsWith(TempCacheRootPath, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Некорректный путь во внутреннем хранилище.");
        }

        return fullPath;
    }

    private static string BuildTempCacheRoot(string workspacePath)
    {
        var hashBytes = SHA256.HashData(Encoding.UTF8.GetBytes(Path.GetFullPath(workspacePath)));
        var hash = Convert.ToHexString(hashBytes[..6]);
        return Path.Combine(Path.GetTempPath(), "PDnDocuments", hash);
    }

    private string TempCacheRootPath => BuildTempCacheRoot(CurrentWorkspacePath);

    private static string NormalizeEntryPath(string path)
    {
        return path
            .Replace('\\', '/')
            .TrimStart('/')
            .Replace("//", "/", StringComparison.Ordinal);
    }

    private static void HidePathIfPossible(string path)
    {
        try
        {
            if (!File.Exists(path) && !Directory.Exists(path))
            {
                return;
            }

            var attributes = File.GetAttributes(path);
            if ((attributes & FileAttributes.Hidden) == 0)
            {
                File.SetAttributes(path, attributes | FileAttributes.Hidden);
            }
        }
        catch
        {
            // Hiding is best-effort only.
        }
    }

    private void NormalizeState(AppState state)
    {
        state.Companies ??= [];
        state.Sections ??= [];
        state.Documents ??= [];
        state.SopdRecords ??= [];
        state.DocumentHistory ??= [];
        state.Sequence ??= new SequenceState();

        state.Sequence.NextCompanyId = Math.Max(state.Sequence.NextCompanyId, NextId(state.Companies.Select(item => item.Id)));
        state.Sequence.NextSectionId = Math.Max(state.Sequence.NextSectionId, NextId(state.Sections.Select(item => item.Id)));
        state.Sequence.NextDocumentId = Math.Max(state.Sequence.NextDocumentId, NextId(state.Documents.Select(item => item.Id)));
        state.Sequence.NextSopdId = Math.Max(state.Sequence.NextSopdId, NextId(state.SopdRecords.Select(item => item.Id)));
        state.Sequence.NextHistoryId = Math.Max(state.Sequence.NextHistoryId, NextId(state.DocumentHistory.Select(item => item.Id)));

        foreach (var company in state.Companies)
        {
            company.Name = (company.Name ?? string.Empty).Trim();
            company.CreatedBy ??= _currentUser;
            company.CreatedAt ??= DateTimeOffset.Now.ToString("O");
        }

        foreach (var section in state.Sections)
        {
            section.Name = (section.Name ?? string.Empty).Trim();
            section.CreatedBy ??= _currentUser;
            section.CreatedAt ??= DateTimeOffset.Now.ToString("O");
        }

        foreach (var document in state.Documents)
        {
            document.Title = (document.Title ?? string.Empty).Trim();
            document.Status = NormalizeStatus(document.Status);
            document.Comment ??= string.Empty;
            document.SectionIds ??= [];
            document.CreatedBy ??= _currentUser;
            document.UpdatedBy ??= _currentUser;
            document.CreatedAt ??= DateTimeOffset.Now.ToString("O");
            document.UpdatedAt ??= document.CreatedAt;
        }

        foreach (var sopd in state.SopdRecords)
        {
            sopd.ConsentType = (sopd.ConsentType ?? string.Empty).Trim();
            sopd.Purpose = (sopd.Purpose ?? string.Empty).Trim();
            sopd.LegalBasis ??= string.Empty;
            sopd.PDCategories ??= string.Empty;
            sopd.DataSubjects ??= string.Empty;
            sopd.PDList ??= string.Empty;
            sopd.ProcessingOperations ??= string.Empty;
            sopd.ProcessingMethod ??= string.Empty;
            sopd.TransferTo ??= string.Empty;
            sopd.Description ??= string.Empty;
            sopd.ValidityPeriod ??= string.Empty;
            sopd.ThirdPartyTransfer = NormalizeTransferOption(sopd.ThirdPartyTransfer);
            sopd.CreatedBy ??= _currentUser;
            sopd.UpdatedBy ??= _currentUser;
            sopd.CreatedAt ??= DateTimeOffset.Now.ToString("O");
            sopd.UpdatedAt ??= sopd.CreatedAt;
        }

        foreach (var historyItem in state.DocumentHistory)
        {
            historyItem.EventText ??= string.Empty;
            historyItem.EventType ??= string.Empty;
            historyItem.CreatedBy ??= _currentUser;
            historyItem.CreatedAt ??= DateTimeOffset.Now.ToString("O");
        }
    }

    private static int NextId(IEnumerable<int> ids) => ids.DefaultIfEmpty(0).Max() + 1;

    private static string NormalizeStatus(string? status) =>
        AppConstants.DocumentStatuses.Contains(status)
            ? status!
            : AppConstants.DefaultStatus;

    private static string NormalizeTransferOption(string? value) =>
        AppConstants.TransferOptions.Contains(value)
            ? value!
            : "Не указано";

    private static string CompanyFolder(int companyId) => $"company-{companyId:D4}";

    private static string DocumentFolder(int documentId) => $"doc-{documentId:D4}";

    private static string SopdFolder(int sopdId) => $"sopd-{sopdId:D4}";

    private static string SanitizeFileName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(value.Where(ch => !invalid.Contains(ch)).ToArray()).Trim();
        return string.IsNullOrWhiteSpace(cleaned) ? "file" : cleaned.Replace(' ', '_');
    }

    private sealed class WorkspaceSettings
    {
        public string WorkspacePath { get; set; } = string.Empty;
    }
}
