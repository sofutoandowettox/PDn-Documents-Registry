using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using PDnDocuments;

namespace PDnDocuments.Wpf;

public partial class MainWindow : Window
{
    private readonly WorkspaceStore _store;
    private AppState _state;
    private readonly Dictionary<int, Border> _documentCardBorders = [];
    private readonly Dictionary<int, Border> _sopdCardBorders = [];

    private bool _isRefreshingFilters;
    private bool _isLoadingEditor;
    private bool _isLoadingSopdEditor;
    private bool _isDocumentDirty;
    private bool _isSopdDirty;
    private bool _isSidebarCollapsed;
    private string _currentPage = "dashboard";
    private int? _selectedDocumentId;
    private int? _selectedSopdId;
    private string? _pendingPdfSourcePath;
    private string? _pendingOfficeSourcePath;
    private string? _pendingSopdAttachmentSourcePath;

    private static readonly Dictionary<string, (string Title, string Subtitle, string PrimaryAction)> PageMeta = new()
    {
        ["dashboard"] = ("Дашборд", string.Empty, "Добавить компанию"),
        ["attention"] = ("Требует внимания", "Просрочки, проблемные карточки и недавние изменения.", "Добавить документ"),
        ["documents"] = ("Документы", string.Empty, "Добавить документ"),
        ["sopd"] = ("СОПД", "Карточки согласий по компаниям и вложениям.", "Добавить карточку"),
        ["settings"] = ("Компании и разделы", "Структура компаний, разделов и рабочего пространства.", "Добавить компанию"),
    };

    public MainWindow()
    {
        InitializeComponent();

        _store = new WorkspaceStore(AppContext.BaseDirectory, App.CurrentUserLabel);
        _state = _store.LoadState();

        SearchTextBox.Text = string.Empty;
        SearchTextBox.ToolTip = "Найти документ, компанию или комментарий";
        ReviewDateTextBox.ToolTip = "Формат: дд.мм.гггг";
        AcceptDateTextBox.ToolTip = "Формат: дд.мм.гггг";

        WireEditorChangeTracking();
        PopulateEditorStatuses();
        PopulateSopdTransferOptions();
        ApplySidebarVisibility();
        RefreshAll();
        SetPage("dashboard");
    }

    protected override void OnClosed(EventArgs e)
    {
        _store.ClearTemporaryFiles();
        base.OnClosed(e);
    }

    private void ApplySidebarVisibility()
    {
        SidebarBorder.Visibility = _isSidebarCollapsed ? Visibility.Collapsed : Visibility.Visible;
        SidebarColumn.Width = _isSidebarCollapsed ? new GridLength(0) : new GridLength(216);
        SidebarGapColumn.Width = _isSidebarCollapsed ? new GridLength(0) : new GridLength(16);
        SidebarToggleButton.ToolTip = _isSidebarCollapsed ? "Показать меню" : "Скрыть меню";
    }

    private void ToggleSidebar()
    {
        _isSidebarCollapsed = !_isSidebarCollapsed;
        ApplySidebarVisibility();
    }

    private void WireEditorChangeTracking()
    {
        DocumentTitleTextBox.TextChanged += (_, _) => MarkDocumentDirty();
        CommentTextBox.TextChanged += (_, _) => MarkDocumentDirty();
        ReviewDateTextBox.TextChanged += (_, _) => MarkDocumentDirty();
        AcceptDateTextBox.TextChanged += (_, _) => MarkDocumentDirty();
        EditorStatusComboBox.SelectionChanged += (_, _) => MarkDocumentDirty();

        SopdCompanyComboBox.SelectionChanged += (_, _) => MarkSopdDirty();
        SopdTitleTextBox.TextChanged += (_, _) => MarkSopdDirty();
        SopdPurposeTextBox.TextChanged += (_, _) => MarkSopdDirty();
        SopdLegalBasisTextBox.TextChanged += (_, _) => MarkSopdDirty();
        SopdCategoriesTextBox.TextChanged += (_, _) => MarkSopdDirty();
        SopdPdListTextBox.TextChanged += (_, _) => MarkSopdDirty();
        SopdSubjectsTextBox.TextChanged += (_, _) => MarkSopdDirty();
        SopdOperationsTextBox.TextChanged += (_, _) => MarkSopdDirty();
        SopdMethodTextBox.TextChanged += (_, _) => MarkSopdDirty();
        SopdTransferComboBox.SelectionChanged += (_, _) => MarkSopdDirty();
        SopdTransferToTextBox.TextChanged += (_, _) => MarkSopdDirty();
        SopdValidityTextBox.TextChanged += (_, _) => MarkSopdDirty();
        SopdDescriptionTextBox.TextChanged += (_, _) => MarkSopdDirty();
    }

    private void MarkDocumentDirty()
    {
        if (_isLoadingEditor)
        {
            return;
        }

        _isDocumentDirty = true;
    }

    private void MarkSopdDirty()
    {
        if (_isLoadingSopdEditor)
        {
            return;
        }

        _isSopdDirty = true;
    }

    private bool MaybeDiscardDocumentChanges()
    {
        if (!_isDocumentDirty)
        {
            return true;
        }

        if (!Confirm("Изменения в карточке документа не сохранены. Продолжить без сохранения?"))
        {
            return false;
        }

        _isDocumentDirty = false;
        return true;
    }

    private bool MaybeDiscardSopdChanges()
    {
        if (!_isSopdDirty)
        {
            return true;
        }

        if (!Confirm("Изменения в карточке СОПД не сохранены. Продолжить без сохранения?"))
        {
            return false;
        }

        _isSopdDirty = false;
        return true;
    }

    private void RefreshAll()
    {
        PopulateFilterCombos();
        PopulateEditorCompanies();
        PopulateSopdCompanies();
        RefreshDashboardPage();
        RefreshAttentionPage();
        RefreshDocumentsPage();
        RefreshSopdPage();
        RefreshSettingsPage();
        RefreshStatusBar();

        if (_selectedDocumentId is int selectedId && _state.Documents.Any(document => document.Id == selectedId))
        {
            LoadDocumentIntoEditor(selectedId);
        }
        else
        {
            ShowRightPlaceholder("Выберите документ слева", "Карточка документа откроется здесь. Справа будут доступны редактирование и файлы.");
        }

        if (_selectedSopdId is int sopdId && _state.SopdRecords.Any(record => record.Id == sopdId))
        {
            LoadSopdIntoEditor(sopdId);
        }
        else
        {
            ClearSopdEditor(CurrentCompanyFilterId());
        }
    }

    private void PopulateEditorStatuses()
    {
        EditorStatusComboBox.Items.Clear();
        foreach (var status in AppConstants.DocumentStatuses)
        {
            EditorStatusComboBox.Items.Add(new OptionItem(status, status));
        }
    }

    private void PopulateSopdTransferOptions()
    {
        SopdTransferComboBox.Items.Clear();
        foreach (var option in AppConstants.TransferOptions)
        {
            SopdTransferComboBox.Items.Add(new OptionItem(option, option));
        }
    }

    private void PopulateFilterCombos()
    {
        _isRefreshingFilters = true;
        try
        {
            var currentCompany = SelectedNullableInt(CompanyFilterComboBox);
            var currentStatus = SelectedString(StatusFilterComboBox);
            var currentSection = SelectedNullableInt(SectionFilterComboBox);
            var currentProblem = SelectedString(ProblemFilterComboBox);

            CompanyFilterComboBox.Items.Clear();
            CompanyFilterComboBox.Items.Add(new OptionItem("Все компании", null));
            foreach (var company in _state.Companies.OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase))
            {
                CompanyFilterComboBox.Items.Add(new OptionItem(company.Name, company.Id));
            }
            SelectComboValue(CompanyFilterComboBox, currentCompany);

            StatusFilterComboBox.Items.Clear();
            StatusFilterComboBox.Items.Add(new OptionItem("Все статусы", null));
            foreach (var status in AppConstants.DocumentStatuses)
            {
                StatusFilterComboBox.Items.Add(new OptionItem(status, status));
            }
            SelectComboValue(StatusFilterComboBox, currentStatus);

            ProblemFilterComboBox.Items.Clear();
            ProblemFilterComboBox.Items.Add(new OptionItem("Все проблемы", "all"));
            ProblemFilterComboBox.Items.Add(new OptionItem("Без PDF", "missing-pdf"));
            ProblemFilterComboBox.Items.Add(new OptionItem("Нужен пересмотр", "due-review"));
            ProblemFilterComboBox.Items.Add(new OptionItem("Скоро пересмотр", "upcoming-review"));
            ProblemFilterComboBox.Items.Add(new OptionItem("Нет даты пересмотра", "missing-review"));
            ProblemFilterComboBox.Items.Add(new OptionItem("Без раздела", "missing-section"));
            ProblemFilterComboBox.Items.Add(new OptionItem("Недавно обновлены", "recent-update"));
            SelectComboValue(ProblemFilterComboBox, currentProblem ?? "all");

            PopulateSectionFilterCombo(currentSection);
        }
        finally
        {
            _isRefreshingFilters = false;
        }
    }

    private void PopulateSectionFilterCombo(int? preferredSectionId)
    {
        var companyId = SelectedNullableInt(CompanyFilterComboBox);
        SectionFilterComboBox.Items.Clear();
        SectionFilterComboBox.Items.Add(new OptionItem("Все разделы", null));

        IEnumerable<SectionRecord> sections = _state.Sections;
        if (companyId is int scopedCompanyId)
        {
            sections = sections.Where(section => section.CompanyId == scopedCompanyId);
        }

        foreach (var section in sections
                     .OrderBy(section => GetCompanyName(section.CompanyId), StringComparer.CurrentCultureIgnoreCase)
                     .ThenBy(section => section.Name, StringComparer.CurrentCultureIgnoreCase))
        {
            var label = companyId is int ? section.Name : $"{GetCompanyName(section.CompanyId)} / {section.Name}";
            SectionFilterComboBox.Items.Add(new OptionItem(label, section.Id));
        }

        SelectComboValue(SectionFilterComboBox, preferredSectionId);
    }

    private void PopulateEditorCompanies()
    {
        var currentCompanyId = SelectedNullableInt(EditorCompanyComboBox) ?? CurrentCompanyFilterId() ?? _state.Companies.OrderBy(item => item.Name).FirstOrDefault()?.Id;

        _isLoadingEditor = true;
        try
        {
            EditorCompanyComboBox.Items.Clear();
            foreach (var company in _state.Companies.OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase))
            {
                EditorCompanyComboBox.Items.Add(new OptionItem(company.Name, company.Id));
            }

            if (EditorCompanyComboBox.Items.Count == 0)
            {
                EditorCompanyComboBox.SelectedIndex = -1;
                PopulateSectionsPanel(null, []);
                return;
            }

            SelectComboValue(EditorCompanyComboBox, currentCompanyId);
            PopulateSectionsPanel(SelectedNullableInt(EditorCompanyComboBox), GetCheckedSectionIds());
        }
        finally
        {
            _isLoadingEditor = false;
        }
    }

    private void PopulateSopdCompanies()
    {
        var currentCompanyId = SelectedNullableInt(SopdCompanyComboBox) ?? CurrentCompanyFilterId() ?? _state.Companies.OrderBy(item => item.Name).FirstOrDefault()?.Id;

        SopdCompanyComboBox.Items.Clear();
        foreach (var company in _state.Companies.OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase))
        {
            SopdCompanyComboBox.Items.Add(new OptionItem(company.Name, company.Id));
        }

        if (SopdCompanyComboBox.Items.Count == 0)
        {
            SopdCompanyComboBox.SelectedIndex = -1;
            return;
        }

        SelectComboValue(SopdCompanyComboBox, currentCompanyId);
    }

    private void RefreshDashboardPage()
    {
        var documents = FilteredDocuments("dashboard").ToList();
        var dueReviews = documents.Count(document => DocumentNeedsAttention(document, "due-review"));
        var missingPdf = documents.Count(document => DocumentNeedsAttention(document, "missing-pdf"));
        var upcomingReviews = documents.Count(document =>
            ReviewPriority(document) is "mid" or "low");

        DashboardDocsMetricTextBlock.Text = documents.Count.ToString(CultureInfo.CurrentCulture);
        DashboardMissingMetricTextBlock.Text = missingPdf.ToString(CultureInfo.CurrentCulture);
        DashboardDueMetricTextBlock.Text = dueReviews.ToString(CultureInfo.CurrentCulture);
        DashboardUpcomingMetricTextBlock.Text = upcomingReviews.ToString(CultureInfo.CurrentCulture);
        DashboardModernDocsMetricTextBlock.Text = DashboardDocsMetricTextBlock.Text;
        DashboardModernMissingMetricTextBlock.Text = DashboardMissingMetricTextBlock.Text;
        DashboardModernDueMetricTextBlock.Text = DashboardDueMetricTextBlock.Text;
        DashboardModernUpcomingMetricTextBlock.Text = DashboardUpcomingMetricTextBlock.Text;
        DashboardModernDocsNoteTextBlock.Text = "Во всех выбранных компаниях и разделах";
        DashboardModernMissingNoteTextBlock.Text = "Документы без основного подписанного PDF";
        DashboardModernDueNoteTextBlock.Text = "Срок наступил или уже истёк";
        DashboardModernUpcomingNoteTextBlock.Text = "Срок пересмотра наступит в ближайшие 30 дней";
        DashboardScopeTextBlock.Text = BuildScopeText("dashboard");
        PopulateDashboardActionBanner(documents);

        var missingReviewDates = documents.Count(document => DocumentNeedsAttention(document, "missing-review"));
        var healthyDocuments = documents.Count(document =>
            !DocumentNeedsAttention(document, "missing-pdf") &&
            !DocumentNeedsAttention(document, "due-review") &&
            !DocumentNeedsAttention(document, "upcoming-review") &&
            !DocumentNeedsAttention(document, "missing-review"));

        UpdateDashboardChart(
            ("Нужен пересмотр", dueReviews, "#F58FA1"),
            ("Без PDF", missingPdf, "#FF7A88"),
            ("Скоро пересмотр", upcomingReviews, "#FFB86C"),
            ("Без даты пересмотра", missingReviewDates, "#7E7891"),
            ("В порядке", healthyDocuments, "#B7AFD6"));

        PopulateDashboardWorkspacePanel(dueReviews, missingPdf, upcomingReviews);
        PopulateDashboardCompaniesPanel(documents);
        PopulateDashboardStatusesPanel(documents);
        PopulateDashboardCoveragePanel(documents);
        PopulateDashboardReviewTimelinePanel(documents);
        PopulateDashboardAttentionPanel();
        PopulateDashboardRecentPanel();
    }

    private void RefreshAttentionPage()
    {
        var rows = FilteredDocuments("attention").ToList();
        var urgent = rows
            .Select(document => new
            {
                Document = document,
                Severity = CalculateAttentionSeverity(document),
                Issues = BuildAttentionLines(document).ToList(),
            })
            .Where(item => item.Severity >= 4 && item.Issues.Count > 0)
            .OrderByDescending(item => item.Severity)
            .ThenBy(item => ReviewPrioritySort(item.Document))
            .ThenBy(item => item.Document.Title, StringComparer.CurrentCultureIgnoreCase)
            .Take(8)
            .ToList();

        var issues = rows
            .Select(document => new
            {
                Document = document,
                Severity = CalculateAttentionSeverity(document),
                Issues = BuildAttentionLines(document).ToList(),
            })
            .Where(item => item.Severity > 0 && item.Severity < 4 && item.Issues.Count > 0)
            .OrderByDescending(item => item.Severity)
            .ThenBy(item => ReviewPrioritySort(item.Document))
            .ThenBy(item => item.Document.Title, StringComparer.CurrentCultureIgnoreCase)
            .Take(10)
            .ToList();

        var upcoming = rows
            .Count(document => ReviewPriority(document) is "mid" or "low");

        var overdue = rows
            .Count(document => ReviewPriority(document) == "high");

        var needsFixCount = rows
            .Count(document =>
            {
                var severity = CalculateAttentionSeverity(document);
                return severity > 0 && severity < 4 && BuildAttentionLines(document).Any();
            });

        var recent = rows
            .Where(document => IsRecentUpdate(document.UpdatedAt))
            .OrderByDescending(document => ParseSortTimestamp(document.UpdatedAt))
            .Take(8)
            .ToList();

        AttentionUrgentSummaryTextBlock.Text = overdue.ToString(CultureInfo.CurrentCulture);
        AttentionIssueSummaryTextBlock.Text = needsFixCount.ToString(CultureInfo.CurrentCulture);
        AttentionRecentSummaryTextBlock.Text = upcoming.ToString(CultureInfo.CurrentCulture);
        PopulateAttentionHeadline(rows, overdue, needsFixCount, upcoming);

        PopulateStackPanel(
            AttentionUrgentPanel,
            urgent.Select(item => CreateDashboardAttentionCard(item.Document, item.Issues)),
            "Сейчас нет срочных документов.");
        PopulateStackPanel(
            AttentionIssuePanel,
            issues.Select(item => CreateDashboardAttentionCard(item.Document, item.Issues)),
            "Здесь появятся документы без раздела, даты пересмотра или других важных данных.");
        PopulateStackPanel(
            AttentionRecentPanel,
            recent.Select(CreateDashboardRecentChangeCard),
            "Недавних изменений пока нет.");

        var visibleIds = rows.Select(document => document.Id).ToHashSet();
        if (_currentPage == "attention" && _selectedDocumentId is int selectedId && !visibleIds.Contains(selectedId))
        {
            ShowRightPlaceholder("Карточка документа", "Выберите документ из списка внимания. Справа можно сразу проверить поля, статус и файлы.");
        }
    }

    private string GetDashboardChartBucket(DocumentRecord document)
    {
        var priority = ReviewPriority(document);
        if (priority == "high")
        {
            return "overdue";
        }

        if (DocumentNeedsAttention(document, "missing-pdf"))
        {
            return "missing-pdf";
        }

        if (priority is "mid" or "low")
        {
            return "upcoming";
        }

        if (TryParseStoredDate(document.ReviewDue) is null)
        {
            return "missing-date";
        }

        return "ok";
    }

    private void UpdateDashboardChart(params (string Label, int Count, string Color)[] buckets)
    {
        var countBlocks = new[]
        {
            DashboardModernChartCount1,
            DashboardModernChartCount2,
            DashboardModernChartCount3,
            DashboardModernChartCount4,
            DashboardModernChartCount5,
        };
        var labelBlocks = new[]
        {
            DashboardModernChartLabel1,
            DashboardModernChartLabel2,
            DashboardModernChartLabel3,
            DashboardModernChartLabel4,
            DashboardModernChartLabel5,
        };
        var bars = new[]
        {
            DashboardModernChartBar1,
            DashboardModernChartBar2,
            DashboardModernChartBar3,
            DashboardModernChartBar4,
            DashboardModernChartBar5,
        };

        var maxValue = Math.Max(1, buckets.Max(item => item.Count));
        const double maxHeight = 212;

        for (var index = 0; index < buckets.Length; index++)
        {
            countBlocks[index].Text = buckets[index].Count.ToString(CultureInfo.CurrentCulture);
            labelBlocks[index].Text = buckets[index].Label;
            var colorBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(buckets[index].Color));
            countBlocks[index].Foreground = colorBrush;
            bars[index].Background = colorBrush;
            bars[index].Height = buckets[index].Count == 0
                ? 0
                : Math.Max(24, Math.Round(maxHeight * buckets[index].Count / (double)maxValue));
            bars[index].Visibility = buckets[index].Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        }
    }

    private void RefreshDocumentsPage()
    {
        var rows = FilteredDocuments("documents").ToList();
        var withoutFiles = rows.Count(document => DocumentNeedsAttention(document, "missing-pdf"));
        var overdue = rows.Count(document => DocumentNeedsAttention(document, "due-review"));
        var withoutSection = rows.Count(document => DocumentNeedsAttention(document, "missing-section"));

        DocumentsCountTextBlock.Text = $"{rows.Count} документов  ·  без PDF: {withoutFiles}  ·  просрочено: {overdue}  ·  без раздела: {withoutSection}";

        DocumentsListPanel.Children.Clear();
        _documentCardBorders.Clear();

        if (rows.Count == 0)
        {
            DocumentsListPanel.Children.Add(CreateEmptyStateCard(
                "Документы не найдены",
                "Сбросьте фильтры или сразу добавьте новый документ.",
                "Добавить документ",
                StartNewDocument));

            if (_currentPage == "documents")
            {
                ShowRightPlaceholder("Документы", "Список пуст или отфильтрован. Можно сразу создать новую карточку.");
            }

            return;
        }

        foreach (var document in rows)
        {
            var card = CreateDocumentCard(document);
            _documentCardBorders[document.Id] = card;
            DocumentsListPanel.Children.Add(card);
        }

        UpdateDocumentCardSelection();

        var visibleIds = rows.Select(document => document.Id).ToHashSet();
        if (_selectedDocumentId is int selectedId && !visibleIds.Contains(selectedId) && _currentPage == "documents")
        {
            ShowRightPlaceholder("Документы", "Выберите документ из списка или сбросьте фильтры.");
        }
    }

    private Border CreateDocumentCard(DocumentRecord document)
    {
        var border = new Border
        {
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#120D1C")),
            BorderBrush = Brush("StrokeBrush"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(22),
            Padding = new Thickness(18, 16, 18, 16),
            Margin = new Thickness(0, 0, 0, 18),
            Cursor = Cursors.Hand,
            Tag = document.Id,
        };
        border.MouseLeftButtonUp += (_, _) => SelectDocument(document.Id);

        var root = new StackPanel();
        border.Child = root;

        var head = new Grid();
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        head.Children.Add(new TextBlock
        {
            Text = document.Title,
            FontSize = 18,
            FontWeight = FontWeights.Bold,
            Foreground = (Brush)Brush("TextBrush"),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 12, 0),
        });
        var statusBadge = CreateStatusBadge(document.Status);
        Grid.SetColumn(statusBadge, 1);
        head.Children.Add(statusBadge);
        root.Children.Add(head);

        var details = new UniformGrid
        {
            Columns = 3,
            Margin = new Thickness(0, 14, 0, 0),
        };
        details.Children.Add(CreateMetaCell("Компания", new TextBlock
        {
            Text = GetCompanyName(document.CompanyId),
            Foreground = (Brush)Brush("TextSoftBrush"),
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 8, 0, 0),
        }));
        details.Children.Add(CreateMetaCell("Пересмотр", new TextBlock
        {
            Text = FormatDisplayDate(document.ReviewDue, "не назначена"),
            Foreground = (Brush)Brush("TextSoftBrush"),
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 8, 0, 0),
        }));
        details.Children.Add(CreateMetaCell("Файлы", CreateFilesWrap(document)));
        root.Children.Add(details);

        var issues = BuildIssueLines(document).ToList();
        root.Children.Add(new TextBlock
        {
            Margin = new Thickness(0, 14, 0, 0),
            Text = issues.Count > 0 ? string.Join(" · ", issues) : "Карточка заполнена, явных проблем не найдено.",
            Foreground = issues.Count > 0 ? (Brush)Brush("TextSoftBrush") : (Brush)Brush("MutedBrush"),
            TextWrapping = TextWrapping.Wrap,
        });

        return border;
    }

    private static Border CreateMetaCell(string label, UIElement valueElement)
    {
        var panel = new StackPanel
        {
            Margin = new Thickness(0, 0, 14, 0),
        };

        panel.Children.Add(new TextBlock
        {
            Text = label,
            Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B8B0C9")),
            FontSize = 13,
        });
        panel.Children.Add(valueElement);

        return new Border
        {
            Background = Brushes.Transparent,
            Child = panel,
        };
    }

    private UIElement CreateStatusBadge(string status)
    {
        var normalizedStatus = string.IsNullOrWhiteSpace(status) ? AppConstants.DefaultStatus : status;
        var (backgroundHex, borderHex, foregroundHex) = normalizedStatus switch
        {
            "На согласовании" => ("#35F7BC68", "#6BF7BC68", "#FFF7E4C0"),
            "Действует" => ("#262EE6A0", "#592EE6A0", "#FFE8FFF6"),
            "На пересмотре" => ("#34FF7A88", "#6BFF7A88", "#FFFFE7EB"),
            "Архив" => ("#1EFFFFFF", "#32FFFFFF", "#FFD0C7E0"),
            _ => ("#30C8AEFF", "#58C8AEFF", "#FFF7F1FF"),
        };

        var badge = new Border
        {
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(backgroundHex)),
            BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(borderHex)),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(11),
            Padding = new Thickness(10, 5, 10, 5),
            HorizontalAlignment = HorizontalAlignment.Left,
            Margin = new Thickness(0, 8, 0, 0),
        };

        badge.Child = new TextBlock
        {
            Text = normalizedStatus,
            Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(foregroundHex)),
            FontSize = 12,
            FontWeight = FontWeights.Bold,
        };

        return badge;
    }

    private WrapPanel CreateFilesWrap(DocumentRecord document)
    {
        var wrap = new WrapPanel
        {
            Margin = new Thickness(0, 8, 0, 0),
        };

        var pdfExists = _store.RelativePathExists(document.PdfPath);
        var pdfTag = CreateFileTag("PDF", "pdf", pdfExists);
        pdfTag.Margin = new Thickness(0, 0, 8, 8);
        wrap.Children.Add(pdfTag);

        var officePath = document.OfficePath;
        if (!string.IsNullOrWhiteSpace(officePath))
        {
            var officeTag = CreateFileTag(OfficeBadgeText(officePath), OfficeVariant(officePath), _store.RelativePathExists(officePath));
            officeTag.Margin = new Thickness(0, 0, 8, 8);
            wrap.Children.Add(officeTag);
        }
        else
        {
            var missingTag = CreateFileTag("Нет файла", "missing", false);
            missingTag.Margin = new Thickness(0, 0, 8, 8);
            wrap.Children.Add(missingTag);
        }

        return wrap;
    }

    private Border CreateFileTag(string text, string variant, bool isReady)
    {
        var border = new Border
        {
            CornerRadius = new CornerRadius(11),
            Padding = new Thickness(10, 5, 10, 5),
            Margin = new Thickness(0, 0, 8, 0),
            BorderThickness = new Thickness(1),
            HorizontalAlignment = HorizontalAlignment.Left,
        };

        ApplyFileTagStyle(border, text, variant, isReady);
        return border;
    }

    private void ApplyFileTagStyle(Border border, string text, string variant, bool isReady)
    {
        Brush background;
        Brush borderBrush;
        Brush foreground;

        if (!isReady || variant == "missing")
        {
            background = Brush("MissingBadgeBrush");
            borderBrush = Brushes.Transparent;
            foreground = Brush("MutedBrush");
        }
        else
        {
            (background, borderBrush, foreground) = variant switch
            {
                "pdf" => (Brush("PdfBadgeBrush"), Brush("PdfBadgeBorderBrush"), Brush("PdfBadgeTextBrush")),
                "excel" => (Brush("ExcelBadgeBrush"), Brush("ExcelBadgeBorderBrush"), Brush("ExcelBadgeTextBrush")),
                _ => (Brush("DocBadgeBrush"), Brush("DocBadgeBorderBrush"), Brush("DocBadgeTextBrush")),
            };
        }

        border.Background = background;
        border.BorderBrush = borderBrush;
        border.Child = new TextBlock
        {
            Text = text,
            Foreground = foreground,
            FontWeight = FontWeights.Bold,
            FontSize = 12,
        };
    }

    private Border CreateEmptyStateCard(string title, string text, string actionText, Action action)
    {
        var border = new Border
        {
            Background = Brush("SurfaceCardBrush"),
            BorderBrush = Brush("StrokeBrush"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(22),
            Padding = new Thickness(24),
        };

        var stack = new StackPanel();
        border.Child = stack;

        stack.Children.Add(new TextBlock
        {
            Text = title,
            FontSize = 20,
            FontWeight = FontWeights.Bold,
            Foreground = (Brush)Brush("TextBrush"),
        });
        stack.Children.Add(new TextBlock
        {
            Text = text,
            Margin = new Thickness(0, 10, 0, 0),
            Foreground = (Brush)Brush("MutedBrush"),
            TextWrapping = TextWrapping.Wrap,
        });

        var button = new Button
        {
            Content = actionText,
            Style = (Style)FindResource("SolidButtonStyle"),
            Margin = new Thickness(0, 18, 0, 0),
            HorizontalAlignment = HorizontalAlignment.Left,
            MinWidth = 180,
        };
        button.Click += (_, _) => action();
        stack.Children.Add(button);

        return border;
    }

    private void PopulateDashboardWorkspacePanel(int dueReviews, int missingPdf, int upcomingReviews)
    {
        DashboardWorkspacePanel.Children.Clear();
        DashboardWorkspacePanel.Children.Add(CreateDashboardInfoRow("Компании", _state.Companies.Count.ToString(CultureInfo.CurrentCulture)));
        DashboardWorkspacePanel.Children.Add(CreateDashboardInfoRow("Разделы", _state.Sections.Count.ToString(CultureInfo.CurrentCulture)));
        DashboardWorkspacePanel.Children.Add(CreateDashboardInfoRow("Документы", _state.Documents.Count.ToString(CultureInfo.CurrentCulture)));
        DashboardWorkspacePanel.Children.Add(CreateDashboardInfoRow("Карточки СОПД", _state.SopdRecords.Count.ToString(CultureInfo.CurrentCulture)));
        DashboardWorkspacePanel.Children.Add(CreateDashboardInfoRow("Без PDF", missingPdf.ToString(CultureInfo.CurrentCulture), highlight: missingPdf > 0));
        DashboardWorkspacePanel.Children.Add(CreateDashboardInfoRow("Нужен пересмотр", dueReviews.ToString(CultureInfo.CurrentCulture), highlight: dueReviews > 0));
        DashboardWorkspacePanel.Children.Add(CreateDashboardInfoRow("Скоро пересмотр", upcomingReviews.ToString(CultureInfo.CurrentCulture), highlight: upcomingReviews > 0));
    }

    private void PopulateDashboardCompaniesPanel(IReadOnlyCollection<DocumentRecord> scopedDocuments)
    {
        DashboardCompaniesPanel.Children.Clear();

        if (_state.Companies.Count == 0)
        {
            DashboardCompaniesPanel.Children.Add(CreateDashboardHint("Пока компаний нет. Добавьте первую компанию из верхней кнопки."));
            return;
        }

        var cards = _state.Companies
            .OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase)
            .Select(company => new
            {
                Company = company,
                Documents = scopedDocuments.Where(document => document.CompanyId == company.Id).ToList(),
            })
            .Where(item => item.Documents.Count > 0 || scopedDocuments.Count == 0)
            .ToList();

        if (cards.Count == 0)
        {
            DashboardCompaniesPanel.Children.Add(CreateDashboardHint("По текущим фильтрам компании с документами не найдены."));
            return;
        }

        foreach (var item in cards)
        {
            DashboardCompaniesPanel.Children.Add(CreateDashboardCompanyCard(item.Company, item.Documents));
        }
    }

    private void PopulateDashboardAttentionPanel()
    {
        DashboardAttentionPanel.Children.Clear();

        var rows = FilteredDocuments("dashboard")
            .Select(document => new
            {
                Document = document,
                Issues = BuildAttentionLines(document).Distinct(StringComparer.CurrentCulture).ToList(),
                Severity = CalculateAttentionSeverity(document),
            })
            .Where(item => item.Issues.Count > 0)
            .OrderByDescending(item => item.Severity)
            .ThenBy(item => item.Document.Title, StringComparer.CurrentCultureIgnoreCase)
            .Take(6)
            .ToList();

        if (rows.Count == 0)
        {
            DashboardAttentionPanel.Children.Add(CreateDashboardHint("Сейчас нет документов с явными проблемами."));
            return;
        }

        foreach (var row in rows)
        {
            DashboardAttentionPanel.Children.Add(CreateDashboardAttentionCard(row.Document, row.Issues));
        }
    }

    private void PopulateDashboardRecentPanel()
    {
        DashboardRecentPanel.Children.Clear();

        var rows = FilteredDocuments("dashboard")
            .Where(document => IsRecentUpdate(document.UpdatedAt))
            .OrderByDescending(document => ParseSortTimestamp(document.UpdatedAt))
            .Take(6)
            .ToList();

        if (rows.Count == 0)
        {
            DashboardRecentPanel.Children.Add(CreateDashboardHint("Недавние изменения появятся после первых обновлений документов."));
            return;
        }

        foreach (var row in rows)
        {
            DashboardRecentPanel.Children.Add(CreateDashboardRecentChangeCard(row));
        }
    }

    private void PopulateDashboardStatusesPanel(IReadOnlyCollection<DocumentRecord> documents)
    {
        DashboardStatusesPanel.Children.Clear();
        if (documents.Count == 0)
        {
            DashboardStatusesPanel.Children.Add(CreateDashboardHint("Статусы появятся после добавления первых документов."));
            return;
        }

        foreach (var status in AppConstants.DocumentStatuses.Where(item => !string.Equals(item, "Архив", StringComparison.Ordinal)))
        {
            var count = documents.Count(document => string.Equals(document.Status, status, StringComparison.Ordinal));
            DashboardStatusesPanel.Children.Add(CreateDashboardProgressCard(
                status,
                count,
                documents.Count,
                StatusColor(status),
                count == 0 ? "Пока нет карточек." : $"{count} из {documents.Count} документов."));
        }
    }

    private void PopulateDashboardCoveragePanel(IReadOnlyCollection<DocumentRecord> documents)
    {
        DashboardCoveragePanel.Children.Clear();
        if (documents.Count == 0)
        {
            DashboardCoveragePanel.Children.Add(CreateDashboardHint("Когда появятся документы, здесь будет видно качество заполнения."));
            return;
        }

        var total = documents.Count;
        var pdfCount = documents.Count(document => _store.RelativePathExists(document.PdfPath));
        var officeCount = documents.Count(document => _store.RelativePathExists(document.OfficePath));
        var sectionCount = documents.Count(document => document.SectionIds.Count > 0);
        var reviewCount = documents.Count(document => TryParseStoredDate(document.ReviewDue) is not null);

        DashboardCoveragePanel.Children.Add(CreateDashboardProgressCard("PDF загружен", pdfCount, total, "#FFF58FA1", $"{Math.Round(pdfCount * 100d / total):0}% карточек с PDF."));
        DashboardCoveragePanel.Children.Add(CreateDashboardProgressCard("Word / Excel", officeCount, total, "#FFC8AEFF", $"{Math.Round(officeCount * 100d / total):0}% карточек с доп. файлом."));
        DashboardCoveragePanel.Children.Add(CreateDashboardProgressCard("Есть раздел", sectionCount, total, "#FF7FD8B8", $"{Math.Round(sectionCount * 100d / total):0}% документов привязаны к разделам."));
        DashboardCoveragePanel.Children.Add(CreateDashboardProgressCard("Есть дата пересмотра", reviewCount, total, "#FFFFD08A", $"{Math.Round(reviewCount * 100d / total):0}% карточек готовы к контролю сроков."));
    }

    private void PopulateDashboardReviewTimelinePanel(IReadOnlyCollection<DocumentRecord> documents)
    {
        DashboardReviewTimelinePanel.Children.Clear();

        var rows = documents
            .Where(document => TryParseStoredDate(document.ReviewDue) is not null)
            .OrderBy(document => TryParseStoredDate(document.ReviewDue))
            .ThenBy(document => document.Title, StringComparer.CurrentCultureIgnoreCase)
            .Take(6)
            .ToList();

        if (rows.Count == 0)
        {
            DashboardReviewTimelinePanel.Children.Add(CreateDashboardHint("После заполнения дат пересмотра здесь появится ближайшая очередь."));
            return;
        }

        foreach (var row in rows)
        {
            DashboardReviewTimelinePanel.Children.Add(CreateDashboardReviewTimelineCard(row));
        }
    }

    private void PopulateDashboardActionBanner(IReadOnlyCollection<DocumentRecord> documents)
    {
        var topDocument = documents
            .OrderByDescending(CalculateAttentionSeverity)
            .ThenBy(ReviewPrioritySort)
            .ThenBy(document => document.Title, StringComparer.CurrentCultureIgnoreCase)
            .FirstOrDefault(document => CalculateAttentionSeverity(document) > 0);

        if (topDocument is null)
        {
            DashboardActionTitleTextBlock.Text = "Сегодня всё спокойно";
            DashboardActionNoteTextBlock.Text = "В выбранной области нет документов с критичными замечаниями. Можно перейти к структуре компаний или добавить новую карточку.";
            return;
        }

        DashboardActionTitleTextBlock.Text = "Что сделать сейчас";
        DashboardActionNoteTextBlock.Text = BuildDashboardActionText(topDocument);
    }

    private void PopulateAttentionHeadline(IReadOnlyCollection<DocumentRecord> rows, int overdue, int needsFixCount, int upcoming)
    {
        if (rows.Count == 0)
        {
            AttentionHeadlineTitleTextBlock.Text = "Сейчас раздел пуст";
            AttentionHeadlineBodyTextBlock.Text = "По текущим фильтрам нет документов, требующих внимания или недавно обновленных.";
            return;
        }

        var topDocument = rows
            .OrderByDescending(CalculateAttentionSeverity)
            .ThenBy(ReviewPrioritySort)
            .ThenBy(document => document.Title, StringComparer.CurrentCultureIgnoreCase)
            .FirstOrDefault();

        AttentionHeadlineTitleTextBlock.Text = overdue > 0
            ? "Есть срочные карточки"
            : upcoming > 0
                ? "Есть документы с близким пересмотром"
                : "Пора пройтись по карточкам";
        AttentionHeadlineBodyTextBlock.Text = topDocument is null
            ? $"Нужен пересмотр: {overdue}. Нужно поправить: {needsFixCount}. Скоро пересмотр: {upcoming}."
            : $"{BuildDashboardActionText(topDocument)} Сейчас в списке: просрочено {overdue}, нужно поправить {needsFixCount}, скоро пересмотр {upcoming}.";
    }

    private string BuildDashboardActionText(DocumentRecord document)
    {
        if (DocumentNeedsAttention(document, "missing-pdf"))
        {
            return $"Для «{document.Title}» нет основного PDF. Откройте карточку и загрузите файл, чтобы снять проблему.";
        }

        return ReviewPriority(document) switch
        {
            "high" when IsReviewDueToday(document) => $"У документа «{document.Title}» срок пересмотра наступил сегодня. Обновите дату и проверьте актуальность вложений.",
            "high" => $"У документа «{document.Title}» просрочен пересмотр. Обновите дату и проверьте актуальность вложений.",
            "mid" or "low" => $"Для «{document.Title}» скоро наступит срок пересмотра. Лучше обновить карточку заранее.",
            _ when DocumentNeedsAttention(document, "missing-section") => $"У документа «{document.Title}» не указан раздел. Добавьте структуру для более чистого реестра.",
            _ => $"Проверьте документ «{document.Title}» и обновите данные при необходимости.",
        };
    }

    private Border CreateDashboardProgressCard(string title, int value, int total, string accentHex, string note)
    {
        var accentBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(accentHex));
        var border = new Border
        {
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0FFFFFFF")),
            BorderBrush = Brush("StrokeBrush"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(16),
            Padding = new Thickness(14, 12, 14, 12),
            Margin = new Thickness(0, 0, 0, 10),
        };

        var stack = new StackPanel();
        var head = new Grid();
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        head.Children.Add(new TextBlock
        {
            Text = title,
            Foreground = Brush("TextBrush"),
            FontWeight = FontWeights.Bold,
        });
        var valueBlock = new TextBlock
        {
            Text = total == 0 ? "0" : $"{value}",
            Foreground = accentBrush,
            FontWeight = FontWeights.Bold,
        };
        Grid.SetColumn(valueBlock, 1);
        head.Children.Add(valueBlock);
        stack.Children.Add(head);

        var barGrid = new Grid
        {
            Height = 10,
            Margin = new Thickness(0, 10, 0, 0),
        };
        barGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(Math.Max(value, 0.01), GridUnitType.Star) });
        barGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(Math.Max(total - value, 0.01), GridUnitType.Star) });
        var trackBorder = new Border
        {
            Background = Brush("ChartTrackBrush"),
            CornerRadius = new CornerRadius(999),
        };
        Grid.SetColumnSpan(trackBorder, 2);
        barGrid.Children.Add(trackBorder);
        if (value > 0)
        {
            barGrid.Children.Add(new Border
            {
                Background = accentBrush,
                CornerRadius = new CornerRadius(999),
            });
        }
        stack.Children.Add(barGrid);

        stack.Children.Add(new TextBlock
        {
            Text = note,
            Margin = new Thickness(0, 10, 0, 0),
            Foreground = Brush("TextSoftBrush"),
            TextWrapping = TextWrapping.Wrap,
        });

        border.Child = stack;
        return border;
    }

    private Border CreateDashboardReviewTimelineCard(DocumentRecord document)
    {
        var priority = ReviewPriority(document);
        var accentHex = priority switch
        {
            "high" => "#FFFF7A88",
            "mid" => "#FFFFB86C",
            "low" => "#FFC8AEFF",
            _ => "#FF7E7891",
        };

        var border = new Border
        {
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0FFFFFFF")),
            BorderBrush = Brush("StrokeBrush"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(16),
            Padding = new Thickness(14, 12, 14, 12),
            Margin = new Thickness(0, 0, 0, 10),
            Cursor = Cursors.Hand,
        };
        border.MouseLeftButtonUp += (_, _) => OpenDocumentFromFeed(document.Id, "documents");

        var stack = new StackPanel();
        var head = new Grid();
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        head.Children.Add(new TextBlock
        {
            Text = document.Title,
            Foreground = Brush("TextBrush"),
            FontWeight = FontWeights.Bold,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 12, 0),
        });
        var dateBlock = new TextBlock
        {
            Text = FormatDisplayDate(document.ReviewDue, "не указана"),
            Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(accentHex)),
            FontWeight = FontWeights.Bold,
        };
        Grid.SetColumn(dateBlock, 1);
        head.Children.Add(dateBlock);
        stack.Children.Add(head);

        stack.Children.Add(new TextBlock
        {
            Text = $"{GetCompanyName(document.CompanyId)} · {BuildTimelineBadgeText(document)}",
            Margin = new Thickness(0, 6, 0, 0),
            Foreground = Brush("TextSoftBrush"),
            TextWrapping = TextWrapping.Wrap,
        });

        border.Child = stack;
        return border;
    }

    private string BuildTimelineBadgeText(DocumentRecord document)
    {
        var days = ReviewDeltaDays(document);
        if (days is null)
        {
            return "без даты";
        }

        if (days.Value < 0)
        {
            return $"просрочено на {Math.Abs(days.Value)} дн.";
        }

        if (days.Value == 0)
        {
            return "срок сегодня";
        }

        return $"через {days.Value} дн.";
    }

    private static string StatusColor(string status) => status switch
    {
        "Черновик" => "#FFC8AEFF",
        "На согласовании" => "#FF76C4FF",
        "Действует" => "#FF7FD8B8",
        "На пересмотре" => "#FFFFD08A",
        "Архив" => "#FF7E7891",
        _ => "#FFC8AEFF",
    };

    private Border CreateDashboardInfoRow(string label, string value, bool wrapValue = false, bool highlight = false)
    {
        var border = new Border
        {
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0FFFFFFF")),
            BorderBrush = Brush("StrokeBrush"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(16),
            Padding = new Thickness(12, 10, 12, 10),
            Margin = new Thickness(0, 0, 0, 8),
        };

        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(132) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        grid.Children.Add(new TextBlock
        {
            Text = label,
            Foreground = Brush("MutedBrush"),
            Margin = new Thickness(0, 2, 12, 0),
        });

        var valueText = new TextBlock
        {
            Text = value,
            TextWrapping = wrapValue ? TextWrapping.Wrap : TextWrapping.NoWrap,
            Foreground = highlight ? Brush("AccentBrush") : Brush("TextSoftBrush"),
            FontWeight = highlight ? FontWeights.Bold : FontWeights.SemiBold,
        };
        Grid.SetColumn(valueText, 1);
        grid.Children.Add(valueText);

        border.Child = grid;
        return border;
    }

    private Border CreateDashboardCompanyCard(CompanyRecord company, IReadOnlyCollection<DocumentRecord> documents)
    {
        var sections = _state.Sections.Count(section => section.CompanyId == company.Id);
        var sopd = _state.SopdRecords.Count(record => record.CompanyId == company.Id);
        var overdue = documents.Count(document => DocumentNeedsAttention(document, "due-review"));
        var missingPdf = documents.Count(document => DocumentNeedsAttention(document, "missing-pdf"));
        var upcoming = documents.Count(document => ReviewPriority(document) is "mid" or "low");
        var pdfUploaded = documents.Count(document => _store.RelativePathExists(document.PdfPath));
        var officeUploaded = documents.Count(document => _store.RelativePathExists(document.OfficePath));
        var totalDocuments = documents.Count;

        var border = new Border
        {
            Background = Brush("SurfaceCardBrush"),
            BorderBrush = Brush("StrokeBrush"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(20),
            Padding = new Thickness(16),
            Margin = new Thickness(0, 0, 12, 12),
            Width = 340,
            Cursor = Cursors.Hand,
        };
        border.MouseLeftButtonUp += (_, _) => OpenCompanyScope(company.Id, "documents");

        var stack = new StackPanel();
        stack.Children.Add(new TextBlock
        {
            Text = company.Name,
            FontSize = 18,
            FontWeight = FontWeights.Bold,
            Foreground = Brush("TextBrush"),
        });

        stack.Children.Add(new TextBlock
        {
            Text = $"{totalDocuments} документов · {sections} разделов · СОПД: {sopd}",
            Margin = new Thickness(0, 6, 0, 0),
            Foreground = Brush("MutedBrush"),
            TextWrapping = TextWrapping.Wrap,
        });

        var metricsGrid = new Grid
        {
            Margin = new Thickness(0, 16, 0, 0),
        };
        metricsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        metricsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(12) });
        metricsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        metricsGrid.Children.Add(CreateDashboardUploadTile("PDF", pdfUploaded, totalDocuments, "#FFF58FA1"));
        var officeTile = CreateDashboardUploadTile("DOCX / XLSX", officeUploaded, totalDocuments, "#FFC8AEFF");
        Grid.SetColumn(officeTile, 2);
        metricsGrid.Children.Add(officeTile);
        stack.Children.Add(metricsGrid);

        stack.Children.Add(new TextBlock
        {
            Text = overdue > 0 || upcoming > 0 || missingPdf > 0
                ? $"Нужен пересмотр: {overdue} · скоро пересмотр: {upcoming} · без PDF: {missingPdf}"
                : "Документы компании сейчас без критичных замечаний.",
            Margin = new Thickness(0, 14, 0, 0),
            Foreground = overdue > 0 || upcoming > 0 || missingPdf > 0 ? Brush("AccentBrush") : Brush("MutedBrush"),
            TextWrapping = TextWrapping.Wrap,
        });

        border.Child = stack;
        return border;
    }

    private Border CreateDashboardUploadTile(string label, int uploaded, int total, string accentHex)
    {
        var ratioText = total == 0
            ? "0 / 0"
            : $"{uploaded} / {total}";
        var noteText = total == 0
            ? "Пока без документов."
            : uploaded == total
                ? "Все файлы на месте."
                : $"Загружено {Math.Round(uploaded * 100d / total):0}%";

        var border = new Border
        {
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0DFFFFFF")),
            BorderBrush = Brush("StrokeBrush"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(16),
            Padding = new Thickness(14, 12, 14, 12),
        };

        var stack = new StackPanel();
        stack.Children.Add(new TextBlock
        {
            Text = label,
            Foreground = Brush("MutedBrush"),
            FontWeight = FontWeights.Bold,
        });
        stack.Children.Add(new TextBlock
        {
            Text = ratioText,
            Margin = new Thickness(0, 12, 0, 0),
            Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(accentHex)),
            FontSize = 26,
            FontWeight = FontWeights.Bold,
        });
        stack.Children.Add(new TextBlock
        {
            Text = noteText,
            Margin = new Thickness(0, 8, 0, 0),
            Foreground = Brush("TextSoftBrush"),
            TextWrapping = TextWrapping.Wrap,
        });

        border.Child = stack;
        return border;
    }

    private Border CreateDashboardAttentionCard(DocumentRecord document, IReadOnlyCollection<string> issues)
    {
        var severity = CalculateAttentionSeverity(document);
        var pdfExists = _store.RelativePathExists(document.PdfPath);
        var officeExists = _store.RelativePathExists(document.OfficePath);
        var border = new Border
        {
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#120D1C")),
            BorderBrush = Brush("StrokeBrush"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(22),
            Padding = new Thickness(18, 16, 18, 16),
            Margin = new Thickness(0, 0, 0, 16),
            Cursor = Cursors.Hand,
        };
        border.MouseLeftButtonUp += (_, _) =>
        {
            OpenDocumentFromFeed(document.Id, _currentPage == "attention" ? "attention" : "documents");
        };

        var stack = new StackPanel();

        var head = new Grid();
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var title = new TextBlock
        {
            Text = document.Title,
            FontWeight = FontWeights.Bold,
            FontSize = 17,
            Foreground = Brush("TextBrush"),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 12, 0),
        };
        head.Children.Add(title);

        var attentionBadge = CreateAttentionBadge(BuildAttentionBadgeText(document), severity);
        Grid.SetColumn(attentionBadge, 1);
        head.Children.Add(attentionBadge);
        stack.Children.Add(head);

        stack.Children.Add(new TextBlock
        {
            Text = $"Компания: {GetCompanyName(document.CompanyId)}",
            Margin = new Thickness(0, 8, 0, 0),
            Foreground = Brush("MutedBrush"),
            TextWrapping = TextWrapping.Wrap,
        });

        var extraIssues = issues.Skip(1).ToList();
        stack.Children.Add(new TextBlock
        {
            Text = extraIssues.Count > 0
                ? string.Join(" · ", extraIssues)
                : issues.FirstOrDefault() ?? "Карточка требует проверки.",
            Margin = new Thickness(0, 10, 0, 0),
            Foreground = Brush("TextSoftBrush"),
            TextWrapping = TextWrapping.Wrap,
        });

        var footer = new Grid
        {
            Margin = new Thickness(0, 14, 0, 0),
        };
        footer.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        footer.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        footer.Children.Add(new TextBlock
        {
            Text = $"Пересмотр: {FormatDisplayDate(document.ReviewDue, "не назначена")}",
            Foreground = Brush("MutedBrush"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 12, 0),
        });
        var filesWrap = new WrapPanel
        {
            HorizontalAlignment = HorizontalAlignment.Right,
        };
        filesWrap.Children.Add(CreateFileTag("PDF", "pdf", pdfExists));
        filesWrap.Children.Add(!string.IsNullOrWhiteSpace(document.OfficePath)
            ? CreateFileTag(OfficeBadgeText(document.OfficePath), OfficeVariant(document.OfficePath), officeExists)
            : CreateFileTag("Нет файла", "missing", false));
        Grid.SetColumn(filesWrap, 1);
        footer.Children.Add(filesWrap);
        stack.Children.Add(footer);

        var actions = new WrapPanel
        {
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 14, 0, 0),
        };
        actions.Children.Add(CreateCardActionButton("Перейти к карточке", () =>
            OpenDocumentFromFeed(document.Id, _currentPage == "attention" ? "attention" : "documents")));
        if (!pdfExists)
        {
            actions.Children.Add(CreateCardActionButton("Загрузить PDF", () =>
                QuickAttachDocumentFile(document.Id, "pdf", _currentPage == "attention" ? "attention" : "documents")));
        }
        else if (ReviewPriority(document) is not null)
        {
            actions.Children.Add(CreateCardActionButton("Обновить дату", () =>
                StartReviewUpdate(document.Id, _currentPage == "attention" ? "attention" : "documents")));
        }
        stack.Children.Add(actions);

        border.Child = stack;
        return border;
    }

    private Border CreateDashboardRecentChangeCard(DocumentRecord document)
    {
        var border = new Border
        {
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#120D1C")),
            BorderBrush = Brush("StrokeBrush"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(18),
            Padding = new Thickness(16, 14, 16, 14),
            Margin = new Thickness(0, 0, 0, 14),
            Cursor = Cursors.Hand,
        };
        border.MouseLeftButtonUp += (_, _) =>
        {
            OpenDocumentFromFeed(document.Id, _currentPage == "attention" ? "attention" : "documents");
        };

        var stack = new StackPanel();
        stack.Children.Add(new TextBlock
        {
            Text = document.Title,
            FontWeight = FontWeights.Bold,
            Foreground = Brush("TextBrush"),
            FontSize = 15,
            TextWrapping = TextWrapping.Wrap,
        });
        stack.Children.Add(new TextBlock
        {
            Text = $"{GetCompanyName(document.CompanyId)} · {FormatTimestamp(document.UpdatedAt)} · {document.UpdatedBy}",
            Margin = new Thickness(0, 6, 0, 0),
            Foreground = Brush("MutedBrush"),
            FontSize = 12,
            TextWrapping = TextWrapping.Wrap,
        });
        stack.Children.Add(new TextBlock
        {
            Text = BuildRecentDocumentCaption(document),
            Margin = new Thickness(0, 8, 0, 0),
            Foreground = Brush("TextSoftBrush"),
            TextWrapping = TextWrapping.Wrap,
        });
        var actions = new WrapPanel
        {
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 12, 0, 0),
        };
        actions.Children.Add(CreateCardActionButton("Перейти к карточке", () =>
            OpenDocumentFromFeed(document.Id, _currentPage == "attention" ? "attention" : "documents")));
        stack.Children.Add(actions);

        border.Child = stack;
        return border;
    }

    private Border CreateAttentionBadge(string text, int severity)
    {
        var (backgroundHex, borderHex, foregroundHex) = severity switch
        {
            >= 4 => ("#32FF6A88", "#64FF6A88", "#FFFFE2E8"),
            >= 2 => ("#30FFB86C", "#5AFFB86C", "#FFFFF0D8"),
            _ => ("#18FFFFFF", "#2EFFFFFF", "#FFE9E2F6"),
        };

        var badge = new Border
        {
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(backgroundHex)),
            BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(borderHex)),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(12),
            Padding = new Thickness(10, 5, 10, 5),
            HorizontalAlignment = HorizontalAlignment.Left,
        };
        badge.Child = new TextBlock
        {
            Text = text,
            Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(foregroundHex)),
            FontSize = 12,
            FontWeight = FontWeights.Bold,
        };
        return badge;
    }

    private string BuildAttentionBadgeText(DocumentRecord document)
    {
        if (DocumentNeedsAttention(document, "missing-pdf"))
        {
            return "PDF не загружен";
        }

        if (ReviewPriority(document) == "high")
        {
            return IsReviewDueToday(document) ? "Пересмотр сегодня" : "Просрочен пересмотр";
        }

        if (ReviewPriority(document) is "mid" or "low")
        {
            return "Скоро пересмотр";
        }

        if (DocumentNeedsAttention(document, "missing-section"))
        {
            return "Без раздела";
        }

        return "Нужна проверка";
    }

    private Button CreateCardActionButton(string text, Action action)
    {
        var button = new Button
        {
            Content = text,
            Style = (Style)FindResource("SoftButtonStyle"),
            Height = 36,
            MinWidth = 132,
            Margin = new Thickness(0, 0, 10, 10),
            Padding = new Thickness(14, 8, 14, 8),
        };
        button.Click += (_, _) => action();
        return button;
    }

    private Border CreateDashboardPill(string text)
    {
        var border = new Border
        {
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0DFFFFFF")),
            BorderBrush = Brush("StrokeBrush"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(12),
            Padding = new Thickness(12, 8, 12, 8),
            Margin = new Thickness(0, 0, 8, 8),
        };

        border.Child = new TextBlock
        {
            Text = text,
            Foreground = Brush("TextSoftBrush"),
            FontSize = 12,
            FontWeight = FontWeights.SemiBold,
        };

        return border;
    }

    private Border CreateDashboardHint(string text)
    {
        var border = new Border
        {
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0FFFFFFF")),
            BorderBrush = Brush("StrokeBrush"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(16),
            Padding = new Thickness(14),
        };

        border.Child = new TextBlock
        {
            Text = text,
            Foreground = Brush("MutedBrush"),
            TextWrapping = TextWrapping.Wrap,
        };

        return border;
    }

    private int CalculateDashboardSeverity(DocumentRecord document)
    {
        var severity = 0;

        if (DocumentNeedsAttention(document, "due-review"))
        {
            severity += 4;
        }

        if (DocumentNeedsAttention(document, "missing-pdf"))
        {
            severity += 3;
        }

        if (DocumentNeedsAttention(document, "missing-section") || TryParseStoredDate(document.ReviewDue) is null)
        {
            severity += 1;
        }

        return severity;
    }

    private static DateTimeOffset ParseSortTimestamp(string? value)
    {
        return DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)
            ? parsed
            : DateTimeOffset.MinValue;
    }

    private void SetPage(string pageKey)
    {
        if (_currentPage != pageKey)
        {
            if (!MaybeDiscardSopdChanges())
            {
                return;
            }

            if (!MaybeDiscardDocumentChanges())
            {
                return;
            }
        }

        _currentPage = pageKey;

        DashboardPage.Visibility = pageKey == "dashboard" ? Visibility.Visible : Visibility.Collapsed;
        AttentionPage.Visibility = pageKey == "attention" ? Visibility.Visible : Visibility.Collapsed;
        DocumentsPage.Visibility = pageKey == "documents" ? Visibility.Visible : Visibility.Collapsed;
        SopdPage.Visibility = pageKey == "sopd" ? Visibility.Visible : Visibility.Collapsed;
        SettingsPage.Visibility = pageKey == "settings" ? Visibility.Visible : Visibility.Collapsed;

        var showToolbar = pageKey is "dashboard" or "attention" or "documents" or "sopd";
        var showDocumentsChrome = pageKey is "documents" or "attention" or "sopd";
        FilterBarBorder.Visibility = showToolbar ? Visibility.Visible : Visibility.Collapsed;
        FilterOptionsGrid.Visibility = pageKey == "sopd" ? Visibility.Collapsed : Visibility.Visible;
        RightPanelBorder.Visibility = showDocumentsChrome ? Visibility.Visible : Visibility.Collapsed;
        RightGapColumn.Width = showDocumentsChrome ? new GridLength(16) : new GridLength(0);
        RightPanelColumn.Width = showDocumentsChrome ? new GridLength(430) : new GridLength(0);

        var documentScopePage = pageKey is "dashboard" or "attention" or "documents";
        StatusFilterComboBox.IsEnabled = documentScopePage;
        SectionFilterComboBox.IsEnabled = documentScopePage;
        ProblemFilterComboBox.IsEnabled = documentScopePage;
        PrimaryActionButton.Visibility = pageKey == "dashboard" ? Visibility.Collapsed : Visibility.Visible;

        if (!PageMeta.TryGetValue(pageKey, out var meta))
        {
            meta = PageMeta["dashboard"];
        }

        PageTitleTextBlock.Text = meta.Title;
        PageSubtitleTextBlock.Text = meta.Subtitle;
        PageSubtitleTextBlock.Visibility = string.IsNullOrWhiteSpace(meta.Subtitle) ? Visibility.Collapsed : Visibility.Visible;
        PrimaryActionButton.Content = meta.PrimaryAction;
        SearchPlaceholderTextBlock.Text = pageKey switch
        {
            "attention" => "Найти документ, проблему, компанию или раздел",
            "sopd" => "Найти карточку СОПД",
            _ => "Найти документ, компанию или раздел",
        };

        DashboardNavButton.Style = (Style)FindResource(pageKey == "dashboard" ? "NavButtonActiveStyle" : "NavButtonStyle");
        AttentionNavButton.Style = (Style)FindResource(pageKey == "attention" ? "NavButtonActiveStyle" : "NavButtonStyle");
        DocumentsNavButton.Style = (Style)FindResource(pageKey == "documents" ? "NavButtonActiveStyle" : "NavButtonStyle");
        SopdNavButton.Style = (Style)FindResource(pageKey == "sopd" ? "NavButtonActiveStyle" : "NavButtonStyle");
        SettingsNavButton.Style = (Style)FindResource(pageKey == "settings" ? "NavButtonActiveStyle" : "NavButtonStyle");

        if (pageKey == "documents")
        {
            RefreshDocumentsPage();
            if (_selectedDocumentId is int selectedDocumentId && _state.Documents.Any(document => document.Id == selectedDocumentId))
            {
                LoadDocumentIntoEditor(selectedDocumentId);
            }
            else
            {
                ShowRightPlaceholder("Выберите документ слева", "Карточка документа откроется здесь. Справа будут доступны редактирование и файлы.");
            }
        }
        else if (pageKey == "dashboard")
        {
            RefreshDashboardPage();
        }
        else if (pageKey == "attention")
        {
            RefreshAttentionPage();
            if (_selectedDocumentId is int selectedDocumentId && _state.Documents.Any(document => document.Id == selectedDocumentId))
            {
                LoadDocumentIntoEditor(selectedDocumentId);
            }
            else
            {
                ShowRightPlaceholder("Карточка документа", "Выберите документ из списка внимания. Справа можно сразу проверить поля, статус и файлы.");
            }
        }
        else if (pageKey == "sopd")
        {
            RefreshSopdPage();
            if (_selectedSopdId is int selectedSopdId && _state.SopdRecords.Any(record => record.Id == selectedSopdId))
            {
                LoadSopdIntoEditor(selectedSopdId);
            }
            else if (_state.Companies.Count == 0)
            {
                ShowRightPlaceholder("СОПД", "Сначала добавьте компанию, чтобы создать карточку и прикрепить файл.");
            }
            else
            {
                ClearSopdEditor(CurrentCompanyFilterId());
            }
        }
        else if (pageKey == "settings")
        {
            RefreshSettingsPage();
        }
    }

    private void ShowRightPlaceholder(string title, string subtitle)
    {
        RightPanelTitleTextBlock.Text = title;
        RightPanelSubtitleTextBlock.Text = subtitle;
        RightPanelSubtitleTextBlock.Visibility = string.IsNullOrWhiteSpace(subtitle) ? Visibility.Collapsed : Visibility.Visible;
        RightPlaceholderPanel.Visibility = Visibility.Visible;
        DocumentEditorScrollViewer.Visibility = Visibility.Collapsed;
        SopdEditorScrollViewer.Visibility = Visibility.Collapsed;
        DocumentActionsPanel.Visibility = Visibility.Collapsed;
        SopdActionsPanel.Visibility = Visibility.Collapsed;
        SaveDocumentButton.IsEnabled = false;
        DeleteDocumentButton.IsEnabled = false;
        SaveSopdButton.IsEnabled = false;
        DeleteSopdButton.IsEnabled = false;
    }

    private void ShowDocumentEditor(string title)
    {
        RightPanelTitleTextBlock.Text = title;
        RightPanelSubtitleTextBlock.Text = string.Empty;
        RightPanelSubtitleTextBlock.Visibility = Visibility.Collapsed;
        RightPlaceholderPanel.Visibility = Visibility.Collapsed;
        DocumentEditorScrollViewer.Visibility = Visibility.Visible;
        SopdEditorScrollViewer.Visibility = Visibility.Collapsed;
        DocumentActionsPanel.Visibility = Visibility.Visible;
        SopdActionsPanel.Visibility = Visibility.Collapsed;
        SaveDocumentButton.IsEnabled = true;
        DeleteDocumentButton.IsEnabled = _selectedDocumentId is not null;
        SaveSopdButton.IsEnabled = false;
        DeleteSopdButton.IsEnabled = false;
    }

    private void ShowSopdEditor(string title, string subtitle = "Поля и файл карточки.")
    {
        RightPanelTitleTextBlock.Text = title;
        RightPanelSubtitleTextBlock.Text = subtitle;
        RightPanelSubtitleTextBlock.Visibility = string.IsNullOrWhiteSpace(subtitle) ? Visibility.Collapsed : Visibility.Visible;
        RightPlaceholderPanel.Visibility = Visibility.Collapsed;
        DocumentEditorScrollViewer.Visibility = Visibility.Collapsed;
        SopdEditorScrollViewer.Visibility = Visibility.Visible;
        DocumentActionsPanel.Visibility = Visibility.Collapsed;
        SopdActionsPanel.Visibility = Visibility.Visible;
        SaveDocumentButton.IsEnabled = false;
        DeleteDocumentButton.IsEnabled = false;
        SaveSopdButton.IsEnabled = true;
        DeleteSopdButton.IsEnabled = _selectedSopdId is not null;
    }

    private void OpenDocumentFromFeed(int documentId, string targetPage)
    {
        if (_currentPage != targetPage)
        {
            SetPage(targetPage);
            if (_currentPage != targetPage)
            {
                return;
            }
        }

        SelectDocument(documentId);
    }

    private void StartReviewUpdate(int documentId, string targetPage)
    {
        OpenDocumentFromFeed(documentId, targetPage);
        if (_selectedDocumentId == documentId)
        {
            ReviewDateTextBox.Focus();
            ReviewDateTextBox.SelectAll();
        }
    }

    private void QuickAttachDocumentFile(int documentId, string kind, string targetPage)
    {
        OpenDocumentFromFeed(documentId, targetPage);
        if (_selectedDocumentId == documentId)
        {
            AttachFile(kind);
        }
    }

    private void StartNewDocument()
    {
        if (!MaybeDiscardDocumentChanges())
        {
            return;
        }

        SetPage("documents");

        var preferredCompanyId = CurrentCompanyFilterId() ?? _state.Companies.OrderBy(company => company.Name, StringComparer.CurrentCultureIgnoreCase).FirstOrDefault()?.Id;
        if (preferredCompanyId is not int companyId)
        {
            ShowWarning("Сначала добавьте компанию.");
            return;
        }

        _selectedDocumentId = null;
        _pendingPdfSourcePath = null;
        _pendingOfficeSourcePath = null;
        _isLoadingEditor = true;
        try
        {
            DocumentTitleTextBox.Text = string.Empty;
            CommentTextBox.Text = string.Empty;
            ReviewDateTextBox.Text = string.Empty;
            AcceptDateTextBox.Text = string.Empty;
            SelectComboValue(EditorCompanyComboBox, companyId);
            SelectComboValue(EditorStatusComboBox, AppConstants.DefaultStatus);
            PopulateSectionsPanel(companyId, []);
            UpdateFilePanels();
        }
        finally
        {
            _isLoadingEditor = false;
        }

        ShowDocumentEditor("Новый документ");
        _isDocumentDirty = false;
        DocumentTitleTextBox.Focus();
        SetActionStatus("Открыта новая карточка документа.");
    }

    private void SelectDocument(int documentId)
    {
        if (_selectedDocumentId != documentId && !MaybeDiscardDocumentChanges())
        {
            return;
        }

        _selectedDocumentId = documentId;
        LoadDocumentIntoEditor(documentId);
        UpdateDocumentCardSelection();
    }

    private void LoadDocumentIntoEditor(int documentId)
    {
        var document = _state.Documents.FirstOrDefault(item => item.Id == documentId);
        if (document is null)
        {
            ShowWarning("Документ не найден.");
            return;
        }

        _pendingPdfSourcePath = null;
        _pendingOfficeSourcePath = null;
        _isLoadingEditor = true;
        try
        {
            DocumentTitleTextBox.Text = document.Title;
            CommentTextBox.Text = document.Comment;
            ReviewDateTextBox.Text = FormatEditorDate(document.ReviewDue);
            AcceptDateTextBox.Text = FormatEditorDate(document.AcceptDate);
            SelectComboValue(EditorCompanyComboBox, document.CompanyId);
            SelectComboValue(EditorStatusComboBox, document.Status);
            PopulateSectionsPanel(document.CompanyId, document.SectionIds);
            UpdateFilePanels(document);
        }
        finally
        {
            _isLoadingEditor = false;
        }

        ShowDocumentEditor(document.Title);
        _isDocumentDirty = false;
    }

    private void PopulateSectionsPanel(int? companyId, IReadOnlyCollection<int> checkedIds)
    {
        SectionsPanel.Children.Clear();

        if (companyId is not int selectedCompanyId)
        {
            SectionsPanel.Children.Add(new TextBlock
            {
                Text = "Сначала выберите компанию.",
                Foreground = (Brush)Brush("MutedBrush"),
                TextWrapping = TextWrapping.Wrap,
            });
            return;
        }

        var sections = _state.Sections
            .Where(section => section.CompanyId == selectedCompanyId)
            .OrderBy(section => section.Name, StringComparer.CurrentCultureIgnoreCase)
            .ToList();

        if (sections.Count == 0)
        {
            SectionsPanel.Children.Add(new TextBlock
            {
                Text = "У компании пока нет разделов.",
                Foreground = (Brush)Brush("MutedBrush"),
                TextWrapping = TextWrapping.Wrap,
            });
            return;
        }

        foreach (var section in sections)
        {
            var checkBox = new CheckBox
            {
                Content = section.Name,
                Tag = section.Id,
                Margin = new Thickness(0, 0, 0, 8),
                Foreground = (Brush)Brush("TextSoftBrush"),
                IsChecked = checkedIds.Contains(section.Id),
            };
            checkBox.Checked += (_, _) => MarkDocumentDirty();
            checkBox.Unchecked += (_, _) => MarkDocumentDirty();
            SectionsPanel.Children.Add(checkBox);
        }
    }

    private List<int> GetCheckedSectionIds()
    {
        return SectionsPanel.Children
            .OfType<CheckBox>()
            .Where(checkBox => checkBox.Tag is int && checkBox.IsChecked == true)
            .Select(checkBox => (int)checkBox.Tag)
            .OrderBy(id => id)
            .ToList();
    }

    private void UpdateFilePanels(DocumentRecord? document = null)
    {
        var pdfPath = _pendingPdfSourcePath;
        var officePath = _pendingOfficeSourcePath;
        var pdfReady = false;
        var officeReady = false;

        if (!string.IsNullOrWhiteSpace(pdfPath) && File.Exists(pdfPath))
        {
            pdfReady = true;
            PdfCaptionTextBlock.Text = $"Выбран файл: {Path.GetFileName(pdfPath)}";
        }
        else if (document is not null && _store.RelativePathExists(document.PdfPath))
        {
            pdfPath = _store.ResolveAbsolutePath(document.PdfPath);
            pdfReady = true;
            PdfCaptionTextBlock.Text = $"Файл: {Path.GetFileName(pdfPath)}";
        }
        else
        {
            PdfCaptionTextBlock.Text = "Файл не загружен";
        }

        if (!string.IsNullOrWhiteSpace(officePath) && File.Exists(officePath))
        {
            officeReady = true;
            OfficeCaptionTextBlock.Text = $"Выбран файл: {Path.GetFileName(officePath)}";
        }
        else if (document is not null && _store.RelativePathExists(document.OfficePath))
        {
            officePath = _store.ResolveAbsolutePath(document.OfficePath);
            officeReady = true;
            OfficeCaptionTextBlock.Text = $"Файл: {Path.GetFileName(officePath)}";
        }
        else
        {
            OfficeCaptionTextBlock.Text = "Файл не загружен";
        }

        SetFileTag(PdfTagBorder, PdfTagTextBlock, "PDF", "pdf", pdfReady);
        var officeVariant = officeReady ? OfficeVariant(officePath) : "missing";
        var officeText = officeReady ? OfficeBadgeText(officePath) : "Нет файла";
        SetFileTag(OfficeTagBorder, OfficeTagTextBlock, officeText, officeVariant, officeReady);

        PdfUploadButton.Content = pdfReady || !string.IsNullOrWhiteSpace(document?.PdfPath) ? "Заменить" : "Загрузить";
        OfficeUploadButton.Content = officeReady || !string.IsNullOrWhiteSpace(document?.OfficePath) ? "Заменить" : "Загрузить";
        PdfOpenButton.IsEnabled = pdfReady;
        PdfDownloadButton.IsEnabled = pdfReady;
        PdfDeleteButton.IsEnabled = pdfReady || (!string.IsNullOrWhiteSpace(document?.PdfPath));
        OfficeOpenButton.IsEnabled = officeReady;
        OfficeDownloadButton.IsEnabled = officeReady;
        OfficeDeleteButton.IsEnabled = officeReady || (!string.IsNullOrWhiteSpace(document?.OfficePath));
    }

    private static bool HasAllowedExtension(string path, params string[] extensions)
    {
        var extension = Path.GetExtension(path);
        return extensions.Any(item => string.Equals(item, extension, StringComparison.OrdinalIgnoreCase));
    }

    private bool TryStageDocumentFile(string kind, string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            ShowWarning("Файл не найден.");
            return false;
        }

        var isValid = kind == "pdf"
            ? HasAllowedExtension(filePath, ".pdf")
            : HasAllowedExtension(filePath, ".doc", ".docx", ".xls", ".xlsx", ".rtf");

        if (!isValid)
        {
            ShowWarning(kind == "pdf"
                ? "Для PDF можно загрузить только файл с расширением .pdf."
                : "Для Office можно загрузить только .doc, .docx, .xls, .xlsx или .rtf.");
            return false;
        }

        if (kind == "pdf")
        {
            _pendingPdfSourcePath = filePath;
        }
        else
        {
            _pendingOfficeSourcePath = filePath;
        }

        UpdateFilePanels(CurrentDocument());
        _isDocumentDirty = true;
        SetActionStatus($"Файл выбран: {Path.GetFileName(filePath)}. Сохраните карточку, чтобы закрепить его в базе.");
        return true;
    }

    private bool TryStageSopdFile(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            ShowWarning("Файл не найден.");
            return false;
        }

        if (!HasAllowedExtension(filePath, ".doc", ".docx", ".rtf"))
        {
            ShowWarning("Для карточки СОПД можно загрузить только .doc, .docx или .rtf.");
            return false;
        }

        _pendingSopdAttachmentSourcePath = filePath;
        UpdateSopdFilePanel(CurrentSopd());
        _isSopdDirty = true;
        SetActionStatus($"Файл выбран: {Path.GetFileName(filePath)}. Сохраните карточку СОПД, чтобы закрепить его в базе.");
        return true;
    }

    private bool TryUploadCompanySopdFile(string filePath)
    {
        var company = SelectedSettingsCompany();
        if (company is null)
        {
            ShowWarning("Сначала выберите компанию.");
            return false;
        }

        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            ShowWarning("Файл не найден.");
            return false;
        }

        if (!HasAllowedExtension(filePath, ".doc", ".docx", ".rtf"))
        {
            ShowWarning("Для общего файла компании можно загрузить только .doc, .docx или .rtf.");
            return false;
        }

        _store.DeleteRelativeFile(company.SopdFilePath);
        company.SopdFilePath = _store.CopyCompanySopdFile(company.Id, filePath);
        PersistState();
        RefreshAll();
        SelectSettingsCompany(company.Id);
        SetActionStatus($"Общий файл компании обновлен: {Path.GetFileName(filePath)}.");
        return true;
    }

    private bool HandleDocumentFileSelection(string kind, string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            ShowWarning("Файл не найден.");
            return false;
        }

        var document = CurrentDocument();
        if (document is null || _isDocumentDirty)
        {
            return TryStageDocumentFile(kind, filePath);
        }

        if (kind == "pdf")
        {
            _store.DeleteRelativeFile(document.PdfPath);
            document.PdfPath = _store.CopyDocumentPdf(document.CompanyId, document.Id, filePath);
            _pendingPdfSourcePath = null;
        }
        else
        {
            _store.DeleteRelativeFile(document.OfficePath);
            document.OfficePath = _store.CopyDocumentOffice(document.CompanyId, document.Id, filePath);
            _pendingOfficeSourcePath = null;
        }

        document.UpdatedAt = NowIso();
        document.UpdatedBy = App.CurrentUserLabel;
        PersistState();
        RefreshAll();
        LoadDocumentIntoEditor(document.Id);
        SetActionStatus($"Файл загружен: {Path.GetFileName(filePath)}.");
        return true;
    }

    private bool HandleSopdFileSelection(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            ShowWarning("Файл не найден.");
            return false;
        }

        var record = CurrentSopd();
        if (record is null || _isSopdDirty)
        {
            return TryStageSopdFile(filePath);
        }

        _store.DeleteRelativeFile(record.AttachmentPath);
        record.AttachmentPath = _store.CopySopdAttachment(record.CompanyId, record.Id, filePath);
        record.UpdatedAt = NowIso();
        record.UpdatedBy = App.CurrentUserLabel;
        _pendingSopdAttachmentSourcePath = null;
        PersistState();
        RefreshAll();
        LoadSopdIntoEditor(record.Id);
        SetActionStatus($"Файл карточки СОПД загружен: {Path.GetFileName(filePath)}.");
        return true;
    }

    private static string? ExtractDroppedFile(DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            return null;
        }

        return e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0
            ? files[0]
            : null;
    }

    private bool CanAcceptDrop(string? tag, string? filePath)
    {
        if (string.IsNullOrWhiteSpace(tag) || string.IsNullOrWhiteSpace(filePath))
        {
            return false;
        }

        return tag switch
        {
            "pdf" => HasAllowedExtension(filePath, ".pdf"),
            "office" => HasAllowedExtension(filePath, ".doc", ".docx", ".xls", ".xlsx", ".rtf"),
            "sopd" or "company-sopd" => HasAllowedExtension(filePath, ".doc", ".docx", ".rtf"),
            _ => false,
        };
    }

    private void FileDropSurface_DragOver(object sender, DragEventArgs e)
    {
        var tag = (sender as FrameworkElement)?.Tag as string;
        var filePath = ExtractDroppedFile(e);
        e.Effects = CanAcceptDrop(tag, filePath) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void FileDropSurface_Drop(object sender, DragEventArgs e)
    {
        var tag = (sender as FrameworkElement)?.Tag as string;
        var filePath = ExtractDroppedFile(e);
        if (!CanAcceptDrop(tag, filePath) || string.IsNullOrWhiteSpace(filePath))
        {
            ShowWarning("Сюда нельзя загрузить этот тип файла.");
            return;
        }

        _ = tag switch
        {
            "pdf" => HandleDocumentFileSelection("pdf", filePath),
            "office" => HandleDocumentFileSelection("office", filePath),
            "sopd" => HandleSopdFileSelection(filePath),
            "company-sopd" => TryUploadCompanySopdFile(filePath),
            _ => false,
        };
    }

    private void SetFileTag(Border border, TextBlock textBlock, string text, string variant, bool isReady)
    {
        textBlock.Text = text;
        ApplyFileTagStyle(border, text, variant, isReady);
    }

    private void UpdateDocumentCardSelection()
    {
        foreach (var pair in _documentCardBorders)
        {
            var isSelected = _selectedDocumentId == pair.Key;
            pair.Value.BorderBrush = isSelected ? Brush("AccentBrush") : Brush("StrokeBrush");
            pair.Value.Background = isSelected
                ? new SolidColorBrush((Color)ColorConverter.ConvertFromString("#181126"))
                : new SolidColorBrush((Color)ColorConverter.ConvertFromString("#120D1C"));
        }
    }

    private void SaveDocument()
    {
        var companyId = SelectedNullableInt(EditorCompanyComboBox);
        if (companyId is not int selectedCompanyId)
        {
            ShowWarning("Выберите компанию для документа.");
            return;
        }

        var title = DocumentTitleTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(title))
        {
            ShowWarning("Укажите название документа.");
            return;
        }

        var status = SelectedString(EditorStatusComboBox) ?? AppConstants.DefaultStatus;
        var reviewDate = ParseEditorDate(ReviewDateTextBox.Text.Trim());
        var acceptDate = ParseEditorDate(AcceptDateTextBox.Text.Trim());
        if (ReviewDateTextBox.Text.Trim().Length > 0 && reviewDate is null)
        {
            ShowWarning("Дата пересмотра должна быть в формате дд.мм.гггг.");
            return;
        }

        if (AcceptDateTextBox.Text.Trim().Length > 0 && acceptDate is null)
        {
            ShowWarning("Дата принятия должна быть в формате дд.мм.гггг.");
            return;
        }

        var isNew = _selectedDocumentId is null;
        var document = isNew
            ? new DocumentRecord
            {
                Id = _state.Sequence.NextDocumentId++,
                CreatedAt = NowIso(),
                CreatedBy = App.CurrentUserLabel,
                UpdatedAt = NowIso(),
                UpdatedBy = App.CurrentUserLabel,
                SortOrder = _state.Documents
                    .Where(item => item.CompanyId == selectedCompanyId)
                    .DefaultIfEmpty()
                    .Max(item => item?.SortOrder ?? 0) + 1,
            }
            : _state.Documents.First(item => item.Id == _selectedDocumentId);
        var previousCompanyId = document.CompanyId;

        document.CompanyId = selectedCompanyId;
        document.Title = title;
        document.Status = status;
        document.Comment = CommentTextBox.Text.Trim();
        document.NeedsOffice = false;
        document.ReviewDue = reviewDate;
        document.AcceptDate = acceptDate;
        document.SectionIds = GetCheckedSectionIds();
        document.UpdatedAt = NowIso();
        document.UpdatedBy = App.CurrentUserLabel;

        if (isNew)
        {
            _state.Documents.Add(document);
        }
        else if (previousCompanyId != selectedCompanyId)
        {
            if (_store.RelativePathExists(document.PdfPath))
            {
                var sourcePath = _store.ResolveAbsolutePath(document.PdfPath);
                if (!string.IsNullOrWhiteSpace(sourcePath))
                {
                    var oldRelativePath = document.PdfPath;
                    document.PdfPath = _store.CopyDocumentPdf(document.CompanyId, document.Id, sourcePath);
                    _store.DeleteRelativeFile(oldRelativePath);
                }
            }

            if (_store.RelativePathExists(document.OfficePath))
            {
                var sourcePath = _store.ResolveAbsolutePath(document.OfficePath);
                if (!string.IsNullOrWhiteSpace(sourcePath))
                {
                    var oldRelativePath = document.OfficePath;
                    document.OfficePath = _store.CopyDocumentOffice(document.CompanyId, document.Id, sourcePath);
                    _store.DeleteRelativeFile(oldRelativePath);
                }
            }
        }

        if (!string.IsNullOrWhiteSpace(_pendingPdfSourcePath))
        {
            _store.DeleteRelativeFile(document.PdfPath);
            document.PdfPath = _store.CopyDocumentPdf(document.CompanyId, document.Id, _pendingPdfSourcePath);
        }

        if (!string.IsNullOrWhiteSpace(_pendingOfficeSourcePath))
        {
            _store.DeleteRelativeFile(document.OfficePath);
            document.OfficePath = _store.CopyDocumentOffice(document.CompanyId, document.Id, _pendingOfficeSourcePath);
        }

        _pendingPdfSourcePath = null;
        _pendingOfficeSourcePath = null;
        _selectedDocumentId = document.Id;
        _isDocumentDirty = false;

        PersistState();
        RefreshAll();
        LoadDocumentIntoEditor(document.Id);
        SetActionStatus(isNew ? "Новый документ сохранён." : "Документ сохранён.");
    }

    private void DeleteDocument()
    {
        if (_selectedDocumentId is not int documentId)
        {
            ShowWarning("Сначала выберите документ.");
            return;
        }

        var document = _state.Documents.FirstOrDefault(item => item.Id == documentId);
        if (document is null)
        {
            ShowWarning("Документ не найден.");
            return;
        }

        if (!Confirm($"Удалить документ «{document.Title}»?"))
        {
            return;
        }

        _store.DeleteRelativeFile(document.PdfPath);
        _store.DeleteRelativeFile(document.OfficePath);
        _state.Documents.RemoveAll(item => item.Id == document.Id);
        _selectedDocumentId = null;
        _pendingPdfSourcePath = null;
        _pendingOfficeSourcePath = null;

        PersistState();
        RefreshAll();
        ShowRightPlaceholder("Документ удалён", "Выберите другую карточку слева или создайте новый документ.");
        SetActionStatus($"Документ «{document.Title}» удалён.");
    }

    private void DuplicateCurrentDocument()
    {
        var document = CurrentDocument();
        if (document is null)
        {
            ShowWarning("Сначала выберите документ для дублирования.");
            return;
        }

        if (_isDocumentDirty &&
            !MessageDialog.ShowConfirm(this, "Дублирование документа", "В дубликат попадут только уже сохранённые изменения. Продолжить?", confirmText: "Продолжить", cancelText: "Отмена"))
        {
            return;
        }

        var copy = new DocumentRecord
        {
            Id = _state.Sequence.NextDocumentId++,
            CompanyId = document.CompanyId,
            Title = BuildDuplicateTitle(document.Title, _state.Documents.Select(item => item.Title)),
            Status = document.Status,
            Comment = document.Comment,
            NeedsOffice = false,
            ReviewDue = document.ReviewDue,
            AcceptDate = document.AcceptDate,
            SortOrder = _state.Documents.Where(item => item.CompanyId == document.CompanyId).DefaultIfEmpty().Max(item => item?.SortOrder ?? 0) + 1,
            SectionIds = [.. document.SectionIds],
            CreatedAt = NowIso(),
            CreatedBy = App.CurrentUserLabel,
            UpdatedAt = NowIso(),
            UpdatedBy = App.CurrentUserLabel,
        };

        if (_store.RelativePathExists(document.PdfPath))
        {
            var sourcePath = _store.ResolveAbsolutePath(document.PdfPath);
            if (!string.IsNullOrWhiteSpace(sourcePath))
            {
                copy.PdfPath = _store.CopyDocumentPdf(copy.CompanyId, copy.Id, sourcePath);
            }
        }

        if (_store.RelativePathExists(document.OfficePath))
        {
            var sourcePath = _store.ResolveAbsolutePath(document.OfficePath);
            if (!string.IsNullOrWhiteSpace(sourcePath))
            {
                copy.OfficePath = _store.CopyDocumentOffice(copy.CompanyId, copy.Id, sourcePath);
            }
        }

        _state.Documents.Add(copy);
        PersistState();
        RefreshAll();
        SetPage("documents");
        SelectDocument(copy.Id);
        SetActionStatus($"Создан дубликат документа «{copy.Title}».");
    }

    #pragma warning disable CS0162
    private void AttachFile(string kind)
    {
        var dialog = new OpenFileDialog
        {
            Filter = kind == "pdf"
                ? "PDF (*.pdf)|*.pdf|Все файлы (*.*)|*.*"
                : "Документы (*.doc;*.docx;*.xls;*.xlsx;*.rtf)|*.doc;*.docx;*.xls;*.xlsx;*.rtf|Все файлы (*.*)|*.*",
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        HandleDocumentFileSelection(kind, dialog.FileName);
        return;

        if (kind == "pdf")
        {
            _pendingPdfSourcePath = dialog.FileName;
        }
        else
        {
            _pendingOfficeSourcePath = dialog.FileName;
        }

        UpdateFilePanels(CurrentDocument());
        _isDocumentDirty = true;
        SetActionStatus("Файл выбран. Сохраните карточку, чтобы закрепить его в базе.");
    }

    private void OpenCurrentFile(string kind)
    {
        var path = kind == "pdf" ? _pendingPdfSourcePath : _pendingOfficeSourcePath;
        if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
        {
            OpenPath(path);
            return;
        }

        var document = CurrentDocument();
        var storedPath = kind == "pdf" ? document?.PdfPath : document?.OfficePath;
        var absolutePath = _store.ResolveAbsolutePath(storedPath);
        if (!string.IsNullOrWhiteSpace(absolutePath) && File.Exists(absolutePath))
        {
            OpenPath(absolutePath);
            return;
        }

        ShowWarning("Файл не найден.");
    }

    private void DownloadCurrentFile(string kind)
    {
        var path = kind == "pdf" ? _pendingPdfSourcePath : _pendingOfficeSourcePath;
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            var document = CurrentDocument();
            var storedPath = kind == "pdf" ? document?.PdfPath : document?.OfficePath;
            path = _store.ResolveAbsolutePath(storedPath);
        }

        DownloadAbsoluteFile(
            path,
            kind == "pdf" ? "Скачать PDF как…" : "Скачать файл Word / Excel как…",
            kind == "pdf" ? "PDF сохранён." : "Файл Word / Excel сохранён.");
    }

    private void DeleteCurrentFile(string kind)
    {
        if (kind == "pdf" && !string.IsNullOrWhiteSpace(_pendingPdfSourcePath))
        {
            _pendingPdfSourcePath = null;
            UpdateFilePanels(CurrentDocument());
            _isDocumentDirty = true;
            return;
        }

        if (kind == "office" && !string.IsNullOrWhiteSpace(_pendingOfficeSourcePath))
        {
            _pendingOfficeSourcePath = null;
            UpdateFilePanels(CurrentDocument());
            _isDocumentDirty = true;
            return;
        }

        var document = CurrentDocument();
        if (document is null)
        {
            return;
        }

        var relativePath = kind == "pdf" ? document.PdfPath : document.OfficePath;
        if (string.IsNullOrWhiteSpace(relativePath))
        {
            return;
        }

        if (!Confirm($"Удалить {(kind == "pdf" ? "PDF" : "Office")} файл из карточки?"))
        {
            return;
        }

        _store.DeleteRelativeFile(relativePath);
        if (kind == "pdf")
        {
            document.PdfPath = null;
        }
        else
        {
            document.OfficePath = null;
        }

        document.UpdatedAt = NowIso();
        document.UpdatedBy = App.CurrentUserLabel;
        PersistState();
        RefreshAll();
        LoadDocumentIntoEditor(document.Id);
        SetActionStatus("Файл удалён.");
    }

    private IEnumerable<SopdRecord> FilteredSopdRecords()
    {
        var search = SearchTextBox.Text.Trim();
        var companyId = CurrentCompanyFilterId();

        IEnumerable<SopdRecord> records = _state.SopdRecords;
        if (companyId is int scopedCompanyId)
        {
            records = records.Where(record => record.CompanyId == scopedCompanyId);
        }

        if (!string.IsNullOrWhiteSpace(search))
        {
            records = records.Where(record =>
                ContainsIgnoreCase(record.ConsentType, search) ||
                ContainsIgnoreCase(record.Purpose, search) ||
                ContainsIgnoreCase(record.PDCategories, search) ||
                ContainsIgnoreCase(record.PDList, search) ||
                ContainsIgnoreCase(record.Description, search) ||
                ContainsIgnoreCase(GetCompanyName(record.CompanyId), search));
        }

        return records
            .OrderBy(record => GetCompanyName(record.CompanyId), StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(record => record.SortOrder)
            .ThenBy(record => record.Id)
            .ToList();
    }

    private void RefreshSopdPage()
    {
        var rows = FilteredSopdRecords().ToList();
        SopdCountTextBlock.Text = $"{rows.Count} карточек";

        SopdListPanel.Children.Clear();
        _sopdCardBorders.Clear();

        if (rows.Count == 0)
        {
            SopdListPanel.Children.Add(CreateEmptyStateCard(
                "Карточек СОПД пока нет",
                "Создайте первую карточку и храните все согласия в одном месте.",
                "Добавить карточку",
                StartNewSopdRecord));
            ClearSopdEditor(CurrentCompanyFilterId());
            return;
        }

        foreach (var record in rows)
        {
            var card = CreateSopdCard(record);
            _sopdCardBorders[record.Id] = card;
            SopdListPanel.Children.Add(card);
        }

        UpdateSopdCardSelection();
        var visibleIds = rows.Select(record => record.Id).ToHashSet();
        if (_selectedSopdId is int selectedId && !visibleIds.Contains(selectedId))
        {
            ClearSopdEditor(CurrentCompanyFilterId());
        }
    }

    private Border CreateSopdCard(SopdRecord record)
    {
        var hasAttachment = _store.RelativePathExists(record.AttachmentPath);
        var border = new Border
        {
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#120D1C")),
            BorderBrush = Brush("StrokeBrush"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(22),
            Padding = new Thickness(18, 16, 18, 16),
            Margin = new Thickness(0, 0, 0, 18),
            Cursor = Cursors.Hand,
            Tag = record.Id,
        };
        border.MouseLeftButtonUp += (_, _) => SelectSopd(record.Id);

        var stack = new StackPanel();

        var head = new Grid();
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        head.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        head.Children.Add(new TextBlock
        {
            Text = string.IsNullOrWhiteSpace(record.ConsentType) ? "Карточка СОПД" : record.ConsentType,
            FontSize = 18,
            FontWeight = FontWeights.Bold,
            Foreground = Brush("TextBrush"),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 12, 0),
        });

        var transferText = $"Передача 3-м лицам: {record.ThirdPartyTransfer}";
        if (string.Equals(record.ThirdPartyTransfer, "Да", StringComparison.Ordinal) && !string.IsNullOrWhiteSpace(record.TransferTo))
        {
            transferText = $"Передача 3-м лицам: {record.TransferTo}";
        }

        var transferBlock = CreateAttentionBadge(
            transferText,
            string.Equals(record.ThirdPartyTransfer, "Да", StringComparison.Ordinal) ? 2 : 1);
        Grid.SetColumn(transferBlock, 1);
        head.Children.Add(transferBlock);
        stack.Children.Add(head);

        var fileTag = hasAttachment
            ? CreateFileTag(OfficeBadgeText(record.AttachmentPath), OfficeVariant(record.AttachmentPath), true)
            : CreateFileTag("Нет файла", "missing", false);
        fileTag.Margin = new Thickness(0, 8, 8, 0);

        var details = new Grid
        {
            Margin = new Thickness(0, 14, 0, 0),
        };
        details.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        details.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(18) });
        details.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        details.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        details.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        details.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var companyCell = CreateMetaCell("Компания", new TextBlock
        {
            Text = GetCompanyName(record.CompanyId),
            Foreground = Brush("TextSoftBrush"),
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 8, 0, 0),
        });
        details.Children.Add(companyCell);

        var fileCell = CreateMetaCell("Файл", fileTag);
        Grid.SetColumn(fileCell, 2);
        details.Children.Add(fileCell);

        var purposeCell = CreateMetaCell("Цель", new TextBlock
        {
            Text = string.IsNullOrWhiteSpace(record.Purpose) ? "не указана" : record.Purpose,
            Foreground = Brush("TextSoftBrush"),
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 8, 0, 0),
            TextWrapping = TextWrapping.Wrap,
        });
        Grid.SetRow(purposeCell, 1);
        details.Children.Add(purposeCell);

        var validityCell = CreateMetaCell("Срок", new TextBlock
        {
            Text = string.IsNullOrWhiteSpace(record.ValidityPeriod) ? "не указан" : record.ValidityPeriod,
            Foreground = Brush("TextSoftBrush"),
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 8, 0, 0),
            TextWrapping = TextWrapping.Wrap,
        });
        Grid.SetRow(validityCell, 1);
        Grid.SetColumn(validityCell, 2);
        details.Children.Add(validityCell);

        var categoriesCell = CreateMetaCell("Категории", new TextBlock
        {
            Text = string.IsNullOrWhiteSpace(record.PDCategories) ? "не заполнены" : record.PDCategories,
            Foreground = string.IsNullOrWhiteSpace(record.PDCategories) ? Brush("MutedBrush") : Brush("TextSoftBrush"),
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 8, 0, 0),
            TextWrapping = TextWrapping.Wrap,
        });
        Grid.SetRow(categoriesCell, 2);
        Grid.SetColumnSpan(categoriesCell, 3);
        details.Children.Add(categoriesCell);
        stack.Children.Add(details);

        border.Child = stack;
        return border;
    }

    private void UpdateSopdCardSelection()
    {
        foreach (var pair in _sopdCardBorders)
        {
            var isSelected = _selectedSopdId == pair.Key;
            pair.Value.BorderBrush = isSelected ? Brush("AccentBrush") : Brush("StrokeBrush");
            pair.Value.Background = isSelected
                ? new SolidColorBrush((Color)ColorConverter.ConvertFromString("#181126"))
                : new SolidColorBrush((Color)ColorConverter.ConvertFromString("#120D1C"));
        }
    }

    private SopdRecord? CurrentSopd()
    {
        return _selectedSopdId is int sopdId
            ? _state.SopdRecords.FirstOrDefault(item => item.Id == sopdId)
            : null;
    }

    private void ClearSopdEditor(int? preferredCompanyId = null)
    {
        _isLoadingSopdEditor = true;
        _selectedSopdId = null;
        _pendingSopdAttachmentSourcePath = null;
        var companyId = preferredCompanyId ?? CurrentCompanyFilterId() ?? _state.Companies.OrderBy(item => item.Name).FirstOrDefault()?.Id;
        try
        {
            SelectComboValue(SopdCompanyComboBox, companyId);
            SopdTitleTextBox.Text = string.Empty;
            SopdPurposeTextBox.Text = string.Empty;
            SopdLegalBasisTextBox.Text = string.Empty;
            SopdCategoriesTextBox.Text = string.Empty;
            SopdPdListTextBox.Text = string.Empty;
            SopdSubjectsTextBox.Text = string.Empty;
            SopdOperationsTextBox.Text = string.Empty;
            SopdMethodTextBox.Text = string.Empty;
            SelectComboValue(SopdTransferComboBox, "Не указано");
            SopdTransferToTextBox.Text = string.Empty;
            SopdValidityTextBox.Text = string.Empty;
            SopdDescriptionTextBox.Text = string.Empty;
        }
        finally
        {
            _isLoadingSopdEditor = false;
        }

        DeleteSopdButton.IsEnabled = false;
        UpdateSopdFilePanel();
        UpdateSopdCardSelection();
        _isSopdDirty = false;
        if (_currentPage == "sopd")
        {
            ShowSopdEditor("Новая карточка СОПД");
        }
    }

    private void StartNewSopdRecord()
    {
        if (!MaybeDiscardSopdChanges())
        {
            return;
        }

        SetPage("sopd");

        if (_state.Companies.Count == 0)
        {
            ShowWarning("Сначала добавьте компанию.");
            return;
        }

        ClearSopdEditor(CurrentCompanyFilterId());
        _isSopdDirty = false;
        SopdTitleTextBox.Focus();
        SetActionStatus("Открыта новая карточка СОПД.");
    }

    private void SelectSopd(int recordId)
    {
        if (_selectedSopdId != recordId && !MaybeDiscardSopdChanges())
        {
            return;
        }

        _selectedSopdId = recordId;
        LoadSopdIntoEditor(recordId);
        UpdateSopdCardSelection();
    }

    private void LoadSopdIntoEditor(int recordId)
    {
        var record = _state.SopdRecords.FirstOrDefault(item => item.Id == recordId);
        if (record is null)
        {
            ShowWarning("Карточка СОПД не найдена.");
            return;
        }

        _pendingSopdAttachmentSourcePath = null;
        _isLoadingSopdEditor = true;
        try
        {
            SelectComboValue(SopdCompanyComboBox, record.CompanyId);
            SopdTitleTextBox.Text = record.ConsentType;
            SopdPurposeTextBox.Text = record.Purpose;
            SopdLegalBasisTextBox.Text = record.LegalBasis;
            SopdCategoriesTextBox.Text = record.PDCategories;
            SopdPdListTextBox.Text = record.PDList;
            SopdSubjectsTextBox.Text = record.DataSubjects;
            SopdOperationsTextBox.Text = record.ProcessingOperations;
            SopdMethodTextBox.Text = record.ProcessingMethod;
            SelectComboValue(SopdTransferComboBox, record.ThirdPartyTransfer);
            SopdTransferToTextBox.Text = record.TransferTo;
            SopdValidityTextBox.Text = record.ValidityPeriod;
            SopdDescriptionTextBox.Text = record.Description;
        }
        finally
        {
            _isLoadingSopdEditor = false;
        }

        DeleteSopdButton.IsEnabled = true;
        UpdateSopdFilePanel(record);
        _isSopdDirty = false;
        if (_currentPage == "sopd")
        {
            ShowSopdEditor(string.IsNullOrWhiteSpace(record.ConsentType) ? "Карточка СОПД" : record.ConsentType);
        }
    }

    private void UpdateSopdFilePanel(SopdRecord? record = null)
    {
        var path = _pendingSopdAttachmentSourcePath;
        var ready = false;

        if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
        {
            ready = true;
            SopdFileCaptionTextBlock.Text = $"Выбран файл: {Path.GetFileName(path)}";
        }
        else if (record is not null && _store.RelativePathExists(record.AttachmentPath))
        {
            path = _store.ResolveAbsolutePath(record.AttachmentPath);
            ready = true;
            SopdFileCaptionTextBlock.Text = $"Файл: {Path.GetFileName(path)}";
        }
        else
        {
            SopdFileCaptionTextBlock.Text = "Файл не загружен";
        }

        var text = ready ? OfficeBadgeText(path) : "Нет файла";
        var variant = ready ? OfficeVariant(path) : "missing";
        SetFileTag(SopdFileTagBorder, SopdFileTagTextBlock, text, variant, ready);

        SopdFileUploadButton.Content = ready || !string.IsNullOrWhiteSpace(record?.AttachmentPath) ? "Заменить" : "Загрузить";
        SopdFileUploadButton.IsEnabled = SelectedNullableInt(SopdCompanyComboBox) is not null;
        SopdFileOpenButton.IsEnabled = ready;
        SopdFileDownloadButton.IsEnabled = ready;
        SopdFileDeleteButton.IsEnabled = ready || (!string.IsNullOrWhiteSpace(record?.AttachmentPath));
    }

    private void SaveSopd()
    {
        var companyId = SelectedNullableInt(SopdCompanyComboBox);
        if (companyId is not int selectedCompanyId)
        {
            ShowWarning("Выберите компанию для карточки СОПД.");
            return;
        }

        var title = SopdTitleTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(title))
        {
            ShowWarning("Укажите название карточки СОПД.");
            return;
        }

        var isNew = _selectedSopdId is null;
        var record = isNew
            ? new SopdRecord
            {
                Id = _state.Sequence.NextSopdId++,
                CreatedAt = NowIso(),
                CreatedBy = App.CurrentUserLabel,
                UpdatedAt = NowIso(),
                UpdatedBy = App.CurrentUserLabel,
                SortOrder = _state.SopdRecords
                    .Where(item => item.CompanyId == selectedCompanyId)
                    .DefaultIfEmpty()
                    .Max(item => item?.SortOrder ?? 0) + 1,
            }
            : _state.SopdRecords.First(item => item.Id == _selectedSopdId);
        var previousCompanyId = record.CompanyId;

        record.CompanyId = selectedCompanyId;
        record.ConsentType = title;
        record.Purpose = SopdPurposeTextBox.Text.Trim();
        record.LegalBasis = SopdLegalBasisTextBox.Text.Trim();
        record.PDCategories = SopdCategoriesTextBox.Text.Trim();
        record.PDList = SopdPdListTextBox.Text.Trim();
        record.DataSubjects = SopdSubjectsTextBox.Text.Trim();
        record.ProcessingOperations = SopdOperationsTextBox.Text.Trim();
        record.ProcessingMethod = SopdMethodTextBox.Text.Trim();
        record.ThirdPartyTransfer = SelectedString(SopdTransferComboBox) ?? "Не указано";
        record.TransferTo = SopdTransferToTextBox.Text.Trim();
        record.ValidityPeriod = SopdValidityTextBox.Text.Trim();
        record.Description = SopdDescriptionTextBox.Text.Trim();
        record.UpdatedAt = NowIso();
        record.UpdatedBy = App.CurrentUserLabel;

        if (isNew)
        {
            _state.SopdRecords.Add(record);
        }
        else if (previousCompanyId != selectedCompanyId && _store.RelativePathExists(record.AttachmentPath))
        {
            var sourcePath = _store.ResolveAbsolutePath(record.AttachmentPath);
            if (!string.IsNullOrWhiteSpace(sourcePath))
            {
                var oldRelativePath = record.AttachmentPath;
                record.AttachmentPath = _store.CopySopdAttachment(record.CompanyId, record.Id, sourcePath);
                _store.DeleteRelativeFile(oldRelativePath);
            }
        }

        if (!string.IsNullOrWhiteSpace(_pendingSopdAttachmentSourcePath))
        {
            _store.DeleteRelativeFile(record.AttachmentPath);
            record.AttachmentPath = _store.CopySopdAttachment(record.CompanyId, record.Id, _pendingSopdAttachmentSourcePath);
        }

        _pendingSopdAttachmentSourcePath = null;
        _selectedSopdId = record.Id;
        _isSopdDirty = false;
        PersistState();
        RefreshAll();
        LoadSopdIntoEditor(record.Id);
        SetActionStatus(isNew ? "Новая карточка СОПД сохранена." : "Карточка СОПД сохранена.");
    }

    private void DeleteSopd()
    {
        if (_selectedSopdId is not int recordId)
        {
            ShowWarning("Сначала выберите карточку СОПД.");
            return;
        }

        var record = _state.SopdRecords.FirstOrDefault(item => item.Id == recordId);
        if (record is null)
        {
            ShowWarning("Карточка СОПД не найдена.");
            return;
        }

        if (!Confirm($"Удалить карточку СОПД «{record.ConsentType}»?"))
        {
            return;
        }

        _store.DeleteRelativeFile(record.AttachmentPath);
        _state.SopdRecords.RemoveAll(item => item.Id == recordId);
        _selectedSopdId = null;
        _pendingSopdAttachmentSourcePath = null;
        PersistState();
        RefreshAll();
        ClearSopdEditor(CurrentCompanyFilterId());
        SetActionStatus($"Карточка СОПД «{record.ConsentType}» удалена.");
    }

    private void DuplicateCurrentSopd()
    {
        var record = CurrentSopd();
        if (record is null)
        {
            ShowWarning("Сначала выберите карточку СОПД для дублирования.");
            return;
        }

        if (_isSopdDirty &&
            !MessageDialog.ShowConfirm(this, "Дублирование карточки СОПД", "В дубликат попадут только уже сохранённые изменения. Продолжить?", confirmText: "Продолжить", cancelText: "Отмена"))
        {
            return;
        }

        var copy = new SopdRecord
        {
            Id = _state.Sequence.NextSopdId++,
            CompanyId = record.CompanyId,
            ConsentType = BuildDuplicateTitle(record.ConsentType, _state.SopdRecords.Select(item => item.ConsentType)),
            Purpose = record.Purpose,
            LegalBasis = record.LegalBasis,
            PDCategories = record.PDCategories,
            DataSubjects = record.DataSubjects,
            PDList = record.PDList,
            ProcessingOperations = record.ProcessingOperations,
            ProcessingMethod = record.ProcessingMethod,
            ThirdPartyTransfer = record.ThirdPartyTransfer,
            TransferTo = record.TransferTo,
            Description = record.Description,
            ValidityPeriod = record.ValidityPeriod,
            SortOrder = _state.SopdRecords.Where(item => item.CompanyId == record.CompanyId).DefaultIfEmpty().Max(item => item?.SortOrder ?? 0) + 1,
            CreatedAt = NowIso(),
            CreatedBy = App.CurrentUserLabel,
            UpdatedAt = NowIso(),
            UpdatedBy = App.CurrentUserLabel,
        };

        if (_store.RelativePathExists(record.AttachmentPath))
        {
            var sourcePath = _store.ResolveAbsolutePath(record.AttachmentPath);
            if (!string.IsNullOrWhiteSpace(sourcePath))
            {
                copy.AttachmentPath = _store.CopySopdAttachment(copy.CompanyId, copy.Id, sourcePath);
            }
        }

        _state.SopdRecords.Add(copy);
        PersistState();
        RefreshAll();
        SetPage("sopd");
        SelectSopd(copy.Id);
        SetActionStatus($"Создан дубликат карточки СОПД «{copy.ConsentType}».");
    }

    private void AttachSopdFile()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Документы (*.doc;*.docx;*.rtf)|*.doc;*.docx;*.rtf|Все файлы (*.*)|*.*",
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        HandleSopdFileSelection(dialog.FileName);
        return;

        _pendingSopdAttachmentSourcePath = dialog.FileName;
        UpdateSopdFilePanel(CurrentSopd());
        _isSopdDirty = true;
        SetActionStatus("Файл выбран. Сохраните карточку СОПД, чтобы закрепить его в базе.");
    }

    private void OpenSopdFile()
    {
        var path = _pendingSopdAttachmentSourcePath;
        if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
        {
            OpenPath(path);
            return;
        }

        var record = CurrentSopd();
        var absolutePath = _store.ResolveAbsolutePath(record?.AttachmentPath);
        if (!string.IsNullOrWhiteSpace(absolutePath) && File.Exists(absolutePath))
        {
            OpenPath(absolutePath);
            return;
        }

        ShowWarning("Файл не найден.");
    }

    private void DownloadSopdFile()
    {
        var path = _pendingSopdAttachmentSourcePath;
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            var record = CurrentSopd();
            path = _store.ResolveAbsolutePath(record?.AttachmentPath);
        }

        DownloadAbsoluteFile(path, "Скачать файл карточки СОПД как…", "Файл карточки СОПД сохранён.");
    }

    private void DeleteSopdFile()
    {
        if (!string.IsNullOrWhiteSpace(_pendingSopdAttachmentSourcePath))
        {
            _pendingSopdAttachmentSourcePath = null;
            UpdateSopdFilePanel(CurrentSopd());
            _isSopdDirty = true;
            return;
        }

        var record = CurrentSopd();
        if (record is null || string.IsNullOrWhiteSpace(record.AttachmentPath))
        {
            return;
        }

        if (!Confirm("Удалить файл из карточки СОПД?"))
        {
            return;
        }

        _store.DeleteRelativeFile(record.AttachmentPath);
        record.AttachmentPath = null;
        record.UpdatedAt = NowIso();
        record.UpdatedBy = App.CurrentUserLabel;
        PersistState();
        RefreshAll();
        LoadSopdIntoEditor(record.Id);
        SetActionStatus("Файл карточки СОПД удалён.");
    }

    private void AddCompany()
    {
        var dialog = new InputDialog("Добавить компанию", "Введите название новой компании:");
        dialog.Owner = this;
        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var name = dialog.Value.Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            ShowWarning("Название компании не может быть пустым.");
            return;
        }

        if (_state.Companies.Any(company => string.Equals(company.Name, name, StringComparison.CurrentCultureIgnoreCase)))
        {
            ShowWarning("Компания с таким названием уже существует.");
            return;
        }

        var company = new CompanyRecord
        {
            Id = _state.Sequence.NextCompanyId++,
            Name = name,
            CreatedAt = NowIso(),
            CreatedBy = App.CurrentUserLabel,
        };

        _state.Companies.Add(company);
        PersistState();
        RefreshAll();
        SelectSettingsCompany(company.Id);
        SetActionStatus($"Компания «{company.Name}» добавлена.");
    }

    private void RefreshSettingsPage()
    {
        var preferredCompanyId = SelectedSettingsCompany()?.Id ?? CurrentCompanyFilterId() ?? _state.Companies.OrderBy(item => item.Name).FirstOrDefault()?.Id;

        SettingsCompanyListBox.Items.Clear();
        SettingsCompaniesCaptionTextBlock.Text = _state.Companies.Count == 0
            ? "Пока компаний нет"
            : $"{_state.Companies.Count} компаний в рабочем пространстве";
        foreach (var company in _state.Companies.OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase))
        {
            var docs = _state.Documents.Count(document => document.CompanyId == company.Id);
            var sopd = _state.SopdRecords.Count(record => record.CompanyId == company.Id);
            SettingsCompanyListBox.Items.Add(new ListBoxItem
            {
                Tag = company.Id,
                Content = CreateSettingsCompanyItemContent(company.Name, docs, sopd),
            });
        }

        if (SettingsCompanyListBox.Items.Count == 0)
        {
            SettingsSectionListBox.Items.Clear();
            SettingsSectionsCaptionTextBlock.Text = "Сначала добавьте компанию";
            RefreshSettingsSummary();
            UpdateCompanySopdFilePanel();
            return;
        }

        SelectSettingsCompany(preferredCompanyId ?? 0);
        RefreshSettingsSections();
        RefreshSettingsSummary();
        UpdateCompanySopdFilePanel();
    }

    private void RefreshSettingsSections()
    {
        SettingsSectionListBox.Items.Clear();
        var company = SelectedSettingsCompany();
        if (company is null)
        {
            SettingsSectionsCaptionTextBlock.Text = "Сначала выберите компанию";
            RefreshSettingsSummary();
            UpdateCompanySopdFilePanel();
            return;
        }

        var sections = _state.Sections
            .Where(section => section.CompanyId == company.Id)
            .OrderBy(section => section.SortOrder)
            .ThenBy(section => section.Name, StringComparer.CurrentCultureIgnoreCase)
            .ToList();

        foreach (var section in sections)
        {
            var docsCount = _state.Documents.Count(document => document.SectionIds.Contains(section.Id));
            SettingsSectionListBox.Items.Add(new ListBoxItem
            {
                Tag = section.Id,
                Content = CreateSettingsSectionItemContent(section.Name, docsCount),
            });
        }

        SettingsSectionsCaptionTextBlock.Text = sections.Count == 0
            ? $"У «{company.Name}» пока нет разделов"
            : $"{sections.Count} разделов у «{company.Name}»";

        if (SettingsSectionListBox.Items.Count > 0)
        {
            SettingsSectionListBox.SelectedIndex = 0;
        }

        RefreshSettingsSummary();
        UpdateCompanySopdFilePanel();
    }

    private void RefreshSettingsSummary()
    {
        SettingsSummaryPanel.Children.Clear();
        SettingsSummaryPanel.Children.Add(CreateDashboardInfoRow("Аккаунт", App.CurrentUserLabel));
        SettingsSummaryPanel.Children.Add(CreateDashboardInfoRow("Рабочая папка", _store.CurrentWorkspacePath, wrapValue: true));
        SettingsSummaryPanel.Children.Add(CreateDashboardInfoRow("Хранилилище", _store.ContainerFilePath, wrapValue: true));

        var company = SelectedSettingsCompany();
        if (company is null)
        {
            SettingsSummaryPanel.Children.Add(CreateDashboardHint("Выберите компанию, чтобы увидеть количество документов, разделов и карточек СОПД."));
            return;
        }

        var sections = _state.Sections.Count(section => section.CompanyId == company.Id);
        var docs = _state.Documents.Count(document => document.CompanyId == company.Id);
        var docsWithoutSection = _state.Documents.Count(document => document.CompanyId == company.Id && document.SectionIds.Count == 0);
        var sopd = _state.SopdRecords.Count(record => record.CompanyId == company.Id);

        SettingsSummaryPanel.Children.Add(CreateDashboardInfoRow("Компания", company.Name));
        SettingsSummaryPanel.Children.Add(CreateDashboardInfoRow("Документы", docs.ToString(CultureInfo.CurrentCulture)));
        SettingsSummaryPanel.Children.Add(CreateDashboardInfoRow("Без раздела", docsWithoutSection.ToString(CultureInfo.CurrentCulture), highlight: docsWithoutSection > 0));
        SettingsSummaryPanel.Children.Add(CreateDashboardInfoRow("Разделы", sections.ToString(CultureInfo.CurrentCulture)));
        SettingsSummaryPanel.Children.Add(CreateDashboardInfoRow("СОПД", sopd.ToString(CultureInfo.CurrentCulture)));
    }

    private UIElement CreateSettingsCompanyItemContent(string name, int documentsCount, int sopdCount)
    {
        var stack = new StackPanel();
        stack.Children.Add(new TextBlock
        {
            Text = name,
            FontWeight = FontWeights.Bold,
            Foreground = Brush("TextBrush"),
            TextWrapping = TextWrapping.Wrap,
        });
        stack.Children.Add(new TextBlock
        {
            Text = $"Документы: {documentsCount} · СОПД: {sopdCount}",
            Margin = new Thickness(0, 4, 0, 0),
            Foreground = Brush("MutedBrush"),
            TextWrapping = TextWrapping.Wrap,
        });
        return stack;
    }

    private UIElement CreateSettingsSectionItemContent(string name, int documentsCount)
    {
        var stack = new StackPanel();
        stack.Children.Add(new TextBlock
        {
            Text = name,
            FontWeight = FontWeights.Bold,
            Foreground = Brush("TextBrush"),
            TextWrapping = TextWrapping.Wrap,
        });
        stack.Children.Add(new TextBlock
        {
            Text = documentsCount == 0 ? "Пока без документов" : $"Документов: {documentsCount}",
            Margin = new Thickness(0, 4, 0, 0),
            Foreground = Brush("MutedBrush"),
            TextWrapping = TextWrapping.Wrap,
        });
        return stack;
    }

    private void UpdateCompanySopdFilePanel()
    {
        var company = SelectedSettingsCompany();
        var ready = company is not null && _store.RelativePathExists(company.SopdFilePath);
        var path = company is null ? null : _store.ResolveAbsolutePath(company.SopdFilePath);

        if (ready && path is not null)
        {
            CompanySopdCaptionTextBlock.Text = $"Файл: {Path.GetFileName(path)}";
        }
        else if (company is null)
        {
            CompanySopdCaptionTextBlock.Text = "Сначала выберите компанию.";
        }
        else
        {
            CompanySopdCaptionTextBlock.Text = "Файл не загружен";
        }

        var tagText = ready ? OfficeBadgeText(path) : "Нет файла";
        var variant = ready ? OfficeVariant(path) : "missing";
        SetFileTag(CompanySopdTagBorder, CompanySopdTagTextBlock, tagText, variant, ready);

        CompanySopdUploadButton.Content = ready || (company is not null && !string.IsNullOrWhiteSpace(company.SopdFilePath)) ? "Заменить" : "Загрузить";
        CompanySopdUploadButton.IsEnabled = company is not null;
        CompanySopdOpenButton.IsEnabled = ready;
        CompanySopdDownloadButton.IsEnabled = ready;
        CompanySopdDeleteButton.IsEnabled = company is not null && !string.IsNullOrWhiteSpace(company.SopdFilePath);
    }

    private void SelectSettingsCompany(int companyId)
    {
        for (var index = 0; index < SettingsCompanyListBox.Items.Count; index++)
        {
            if (SettingsCompanyListBox.Items[index] is ListBoxItem item && item.Tag is int currentId && currentId == companyId)
            {
                SettingsCompanyListBox.SelectedIndex = index;
                SettingsCompanyListBox.ScrollIntoView(item);
                return;
            }
        }

        if (SettingsCompanyListBox.Items.Count > 0)
        {
            SettingsCompanyListBox.SelectedIndex = 0;
        }
    }

    private CompanyRecord? SelectedSettingsCompany()
    {
        return SettingsCompanyListBox.SelectedItem is ListBoxItem item && item.Tag is int companyId
            ? _state.Companies.FirstOrDefault(company => company.Id == companyId)
            : null;
    }

    private SectionRecord? SelectedSettingsSection()
    {
        return SettingsSectionListBox.SelectedItem is ListBoxItem item && item.Tag is int sectionId
            ? _state.Sections.FirstOrDefault(section => section.Id == sectionId)
            : null;
    }

    private string UniqueCopyCompanyName(string baseName)
    {
        var index = 1;
        var candidate = $"{baseName} копия";
        while (_state.Companies.Any(company => string.Equals(company.Name, candidate, StringComparison.CurrentCultureIgnoreCase)))
        {
            index++;
            candidate = $"{baseName} копия {index}";
        }

        return candidate;
    }

    private static string BuildDuplicateTitle(string baseName, IEnumerable<string> existingNames)
    {
        var seed = string.IsNullOrWhiteSpace(baseName) ? "Новая карточка" : baseName.Trim();
        var existing = existingNames
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .ToHashSet(StringComparer.CurrentCultureIgnoreCase);

        var index = 1;
        var candidate = $"{seed} копия";
        while (existing.Contains(candidate))
        {
            index++;
            candidate = $"{seed} копия {index}";
        }

        return candidate;
    }

    private CompanyRecord CloneCompanyStructure(int sourceCompanyId, string newName)
    {
        var newCompany = new CompanyRecord
        {
            Id = _state.Sequence.NextCompanyId++,
            Name = newName.Trim(),
            CreatedAt = NowIso(),
            CreatedBy = App.CurrentUserLabel,
        };
        _state.Companies.Add(newCompany);

        var sectionMap = new Dictionary<int, int>();
        foreach (var section in _state.Sections
                     .Where(item => item.CompanyId == sourceCompanyId)
                     .OrderBy(item => item.SortOrder)
                     .ThenBy(item => item.Id))
        {
            var copy = new SectionRecord
            {
                Id = _state.Sequence.NextSectionId++,
                CompanyId = newCompany.Id,
                Name = section.Name,
                SortOrder = section.SortOrder,
                CreatedAt = NowIso(),
                CreatedBy = App.CurrentUserLabel,
            };
            _state.Sections.Add(copy);
            sectionMap[section.Id] = copy.Id;
        }

        var sourceCompany = _state.Companies.FirstOrDefault(item => item.Id == sourceCompanyId);
        if (sourceCompany is not null && _store.RelativePathExists(sourceCompany.SopdFilePath))
        {
            var sourcePath = _store.ResolveAbsolutePath(sourceCompany.SopdFilePath);
            if (!string.IsNullOrWhiteSpace(sourcePath))
            {
                newCompany.SopdFilePath = _store.CopyCompanySopdFile(newCompany.Id, sourcePath);
            }
        }

        foreach (var document in _state.Documents
                     .Where(item => item.CompanyId == sourceCompanyId)
                     .OrderBy(item => item.SortOrder)
                     .ThenBy(item => item.Id)
                     .ToList())
        {
            var copy = new DocumentRecord
            {
                Id = _state.Sequence.NextDocumentId++,
                CompanyId = newCompany.Id,
                Title = document.Title,
                Status = document.Status,
                Comment = document.Comment,
                NeedsOffice = false,
                ReviewDue = document.ReviewDue,
                AcceptDate = document.AcceptDate,
                SortOrder = document.SortOrder,
                SectionIds = document.SectionIds.Where(sectionMap.ContainsKey).Select(id => sectionMap[id]).ToList(),
                CreatedAt = NowIso(),
                CreatedBy = App.CurrentUserLabel,
                UpdatedAt = NowIso(),
                UpdatedBy = App.CurrentUserLabel,
            };

            if (_store.RelativePathExists(document.PdfPath))
            {
                var sourcePath = _store.ResolveAbsolutePath(document.PdfPath);
                if (!string.IsNullOrWhiteSpace(sourcePath))
                {
                    copy.PdfPath = _store.CopyDocumentPdf(newCompany.Id, copy.Id, sourcePath);
                }
            }

            if (_store.RelativePathExists(document.OfficePath))
            {
                var sourcePath = _store.ResolveAbsolutePath(document.OfficePath);
                if (!string.IsNullOrWhiteSpace(sourcePath))
                {
                    copy.OfficePath = _store.CopyDocumentOffice(newCompany.Id, copy.Id, sourcePath);
                }
            }

            _state.Documents.Add(copy);
        }

        foreach (var record in _state.SopdRecords
                     .Where(item => item.CompanyId == sourceCompanyId)
                     .OrderBy(item => item.SortOrder)
                     .ThenBy(item => item.Id)
                     .ToList())
        {
            var copy = new SopdRecord
            {
                Id = _state.Sequence.NextSopdId++,
                CompanyId = newCompany.Id,
                ConsentType = record.ConsentType,
                Purpose = record.Purpose,
                LegalBasis = record.LegalBasis,
                PDCategories = record.PDCategories,
                DataSubjects = record.DataSubjects,
                PDList = record.PDList,
                ProcessingOperations = record.ProcessingOperations,
                ProcessingMethod = record.ProcessingMethod,
                ThirdPartyTransfer = record.ThirdPartyTransfer,
                TransferTo = record.TransferTo,
                Description = record.Description,
                ValidityPeriod = record.ValidityPeriod,
                SortOrder = record.SortOrder,
                CreatedAt = NowIso(),
                CreatedBy = App.CurrentUserLabel,
                UpdatedAt = NowIso(),
                UpdatedBy = App.CurrentUserLabel,
            };

            if (_store.RelativePathExists(record.AttachmentPath))
            {
                var sourcePath = _store.ResolveAbsolutePath(record.AttachmentPath);
                if (!string.IsNullOrWhiteSpace(sourcePath))
                {
                    copy.AttachmentPath = _store.CopySopdAttachment(newCompany.Id, copy.Id, sourcePath);
                }
            }

            _state.SopdRecords.Add(copy);
        }

        return newCompany;
    }

    private void DeleteCompanyStructure(int companyId)
    {
        var company = _state.Companies.FirstOrDefault(item => item.Id == companyId);
        if (company is null)
        {
            return;
        }

        _store.DeleteRelativeFile(company.SopdFilePath);

        foreach (var document in _state.Documents.Where(item => item.CompanyId == companyId).ToList())
        {
            _store.DeleteRelativeFile(document.PdfPath);
            _store.DeleteRelativeFile(document.OfficePath);
        }

        foreach (var record in _state.SopdRecords.Where(item => item.CompanyId == companyId).ToList())
        {
            _store.DeleteRelativeFile(record.AttachmentPath);
        }

        _state.Documents.RemoveAll(item => item.CompanyId == companyId);
        _state.SopdRecords.RemoveAll(item => item.CompanyId == companyId);
        _state.Sections.RemoveAll(item => item.CompanyId == companyId);
        _state.Companies.RemoveAll(item => item.Id == companyId);
        _selectedDocumentId = null;
        _selectedSopdId = null;
    }

    private void RenameSelectedCompany()
    {
        var company = SelectedSettingsCompany();
        if (company is null)
        {
            ShowWarning("Сначала выберите компанию.");
            return;
        }

        var dialog = new InputDialog("Переименовать компанию", "Новое название:", company.Name)
        {
            Owner = this,
        };
        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var name = dialog.Value.Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            ShowWarning("Название компании не может быть пустым.");
            return;
        }

        if (_state.Companies.Any(item => item.Id != company.Id && string.Equals(item.Name, name, StringComparison.CurrentCultureIgnoreCase)))
        {
            ShowWarning("Компания с таким названием уже существует.");
            return;
        }

        company.Name = name;
        PersistState();
        RefreshAll();
        SelectSettingsCompany(company.Id);
        SetActionStatus($"Компания «{company.Name}» переименована.");
    }

    private void CopySelectedCompany()
    {
        var company = SelectedSettingsCompany();
        if (company is null)
        {
            ShowWarning("Сначала выберите компанию.");
            return;
        }

        var dialog = new InputDialog("Копировать компанию", "Название копии:", UniqueCopyCompanyName(company.Name))
        {
            Owner = this,
        };
        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var name = dialog.Value.Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            ShowWarning("Название компании не может быть пустым.");
            return;
        }

        if (_state.Companies.Any(item => string.Equals(item.Name, name, StringComparison.CurrentCultureIgnoreCase)))
        {
            ShowWarning("Компания с таким названием уже существует.");
            return;
        }

        var copy = CloneCompanyStructure(company.Id, name);
        PersistState();
        RefreshAll();
        SelectSettingsCompany(copy.Id);
        SetActionStatus($"Компания «{company.Name}» скопирована.");
    }

    private void DeleteSelectedCompany()
    {
        var company = SelectedSettingsCompany();
        if (company is null)
        {
            ShowWarning("Сначала выберите компанию.");
            return;
        }

        if (!Confirm($"Удалить компанию «{company.Name}» и все её документы и карточки СОПД?"))
        {
            return;
        }

        DeleteCompanyStructure(company.Id);
        PersistState();
        RefreshAll();
        SetActionStatus($"Компания «{company.Name}» удалена.");
    }

    private void AddSectionForSelectedCompany()
    {
        var company = SelectedSettingsCompany();
        if (company is null)
        {
            ShowWarning("Сначала выберите компанию для нового раздела.");
            return;
        }

        var dialog = new InputDialog("Новый раздел", $"Название раздела для «{company.Name}»:")
        {
            Owner = this,
        };
        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var name = dialog.Value.Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            ShowWarning("Название раздела не может быть пустым.");
            return;
        }

        if (_state.Sections.Any(section => section.CompanyId == company.Id && string.Equals(section.Name, name, StringComparison.CurrentCultureIgnoreCase)))
        {
            ShowWarning("Раздел с таким названием уже существует.");
            return;
        }

        _state.Sections.Add(new SectionRecord
        {
            Id = _state.Sequence.NextSectionId++,
            CompanyId = company.Id,
            Name = name,
            SortOrder = _state.Sections.Where(section => section.CompanyId == company.Id).DefaultIfEmpty().Max(section => section?.SortOrder ?? 0) + 1,
            CreatedAt = NowIso(),
            CreatedBy = App.CurrentUserLabel,
        });
        PersistState();
        RefreshAll();
        SelectSettingsCompany(company.Id);
        SetActionStatus($"Раздел «{name}» добавлен.");
    }

    private void RenameSelectedSection()
    {
        var section = SelectedSettingsSection();
        var company = SelectedSettingsCompany();
        if (section is null || company is null)
        {
            ShowWarning("Сначала выберите раздел.");
            return;
        }

        var dialog = new InputDialog("Переименовать раздел", "Новое название:", section.Name)
        {
            Owner = this,
        };
        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var name = dialog.Value.Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            ShowWarning("Название раздела не может быть пустым.");
            return;
        }

        if (_state.Sections.Any(item => item.Id != section.Id && item.CompanyId == company.Id && string.Equals(item.Name, name, StringComparison.CurrentCultureIgnoreCase)))
        {
            ShowWarning("Раздел с таким названием уже существует.");
            return;
        }

        section.Name = name;
        PersistState();
        RefreshAll();
        SelectSettingsCompany(company.Id);
        SetActionStatus($"Раздел «{name}» переименован.");
    }

    private void DeleteSelectedSection()
    {
        var section = SelectedSettingsSection();
        var company = SelectedSettingsCompany();
        if (section is null || company is null)
        {
            ShowWarning("Сначала выберите раздел.");
            return;
        }

        if (!Confirm($"Удалить раздел «{section.Name}»? Документы останутся, исчезнет только привязка к разделу."))
        {
            return;
        }

        foreach (var document in _state.Documents)
        {
            document.SectionIds.RemoveAll(id => id == section.Id);
        }

        _state.Sections.RemoveAll(item => item.Id == section.Id);
        PersistState();
        RefreshAll();
        SelectSettingsCompany(company.Id);
        SetActionStatus($"Раздел «{section.Name}» удалён.");
    }

    private void UploadCompanySopdFile()
    {
        var company = SelectedSettingsCompany();
        if (company is null)
        {
            ShowWarning("Сначала выберите компанию.");
            return;
        }

        var dialog = new OpenFileDialog
        {
            Filter = "Документы (*.doc;*.docx;*.rtf)|*.doc;*.docx;*.rtf|Все файлы (*.*)|*.*",
        };
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        TryUploadCompanySopdFile(dialog.FileName);
        return;

        _store.DeleteRelativeFile(company.SopdFilePath);
        company.SopdFilePath = _store.CopyCompanySopdFile(company.Id, dialog.FileName);
        PersistState();
        RefreshAll();
        SelectSettingsCompany(company.Id);
        SetActionStatus("Общий файл СОПД компании обновлён.");
    }

    private void OpenCompanySopdFile()
    {
        var company = SelectedSettingsCompany();
        var path = _store.ResolveAbsolutePath(company?.SopdFilePath);
        if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
        {
            OpenPath(path);
            return;
        }

        ShowWarning("Файл не найден.");
    }

    private void DownloadCompanySopdFile()
    {
        var company = SelectedSettingsCompany();
        var path = _store.ResolveAbsolutePath(company?.SopdFilePath);
        DownloadAbsoluteFile(path, "Скачать общий файл СОПД как…", "Общий файл СОПД сохранён.");
    }

    private void DeleteCompanySopdFile()
    {
        var company = SelectedSettingsCompany();
        if (company is null || string.IsNullOrWhiteSpace(company.SopdFilePath))
        {
            return;
        }

        if (!Confirm("Удалить общий файл СОПД компании?"))
        {
            return;
        }

        _store.DeleteRelativeFile(company.SopdFilePath);
        company.SopdFilePath = null;
        PersistState();
        RefreshAll();
        SelectSettingsCompany(company.Id);
        SetActionStatus("Общий файл СОПД компании удалён.");
    }

    #pragma warning restore CS0162
    private void ExportDocumentsToCsv()
    {
        var dialog = new SaveFileDialog
        {
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"documents-{DateTime.Now:yyyyMMdd_HHmmss}.csv",
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        var builder = new StringBuilder();
        builder.AppendLine("Название,Компания,Разделы,Статус,Дата_пересмотра,Дата_принятия,PDF,Office,Комментарий,Обновлено");
        foreach (var document in FilteredDocuments("documents"))
        {
            var sections = string.Join(" | ", SectionNames(document.SectionIds));
            builder.AppendLine(string.Join(",",
                Csv(document.Title),
                Csv(GetCompanyName(document.CompanyId)),
                Csv(sections),
                Csv(document.Status),
                Csv(FormatDisplayDate(document.ReviewDue, string.Empty)),
                Csv(FormatDisplayDate(document.AcceptDate, string.Empty)),
                Csv(_store.ResolveAbsolutePath(document.PdfPath) ?? string.Empty),
                Csv(_store.ResolveAbsolutePath(document.OfficePath) ?? string.Empty),
                Csv(document.Comment),
                Csv(FormatTimestamp(document.UpdatedAt))));
        }

        File.WriteAllText(dialog.FileName, builder.ToString(), Encoding.UTF8);
        SetActionStatus($"CSV сохранен: {dialog.FileName}");
    }

    private void CreateBackup()
    {
        if ((_isDocumentDirty || _isSopdDirty) &&
            !MessageDialog.ShowConfirm(this, "Резервная копия", "В резервную копию попадут только уже сохранённые изменения. Продолжить?", confirmText: "Продолжить", cancelText: "Отмена"))
        {
            return;
        }

        var dialog = new SaveFileDialog
        {
            Title = "Создать резервную копию",
            Filter = "Резервная копия (*.pddoc-backup)|*.pddoc-backup|Хранилище (*.pddoc-store)|*.pddoc-store|Все файлы (*.*)|*.*",
            FileName = $"pddoc-backup-{DateTime.Now:yyyyMMdd_HHmmss}.pddoc-backup",
        };
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            _store.CreateBackup(dialog.FileName);
            SetActionStatus($"Резервная копия создана: {dialog.FileName}");
        }
        catch (Exception ex)
        {
            ShowWarning($"Не удалось создать резервную копию: {ex.Message}");
        }
    }

    private void RestoreBackup()
    {
        if (!MaybeDiscardSopdChanges() || !MaybeDiscardDocumentChanges())
        {
            return;
        }

        var dialog = new OpenFileDialog
        {
            Title = "Восстановить из резервной копии",
            Filter = "Резервная копия (*.pddoc-backup;*.pddoc-store)|*.pddoc-backup;*.pddoc-store|Все файлы (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false,
        };
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        if (!MessageDialog.ShowConfirm(this, "Восстановление", "Текущее хранилище будет полностью заменено выбранной копией. Продолжить?", confirmText: "Восстановить", cancelText: "Отмена", useDangerAccent: true))
        {
            return;
        }

        try
        {
            _store.RestoreBackup(dialog.FileName);
            _state = _store.LoadState();
            _selectedDocumentId = null;
            _selectedSopdId = null;
            _pendingPdfSourcePath = null;
            _pendingOfficeSourcePath = null;
            _pendingSopdAttachmentSourcePath = null;
            _isDocumentDirty = false;
            _isSopdDirty = false;
            RefreshAll();
            SetPage("dashboard");
            SetActionStatus($"Хранилище восстановлено из копии: {dialog.FileName}");
        }
        catch (Exception ex)
        {
            ShowWarning($"Не удалось восстановить хранилище: {ex.Message}");
        }
    }

    private void ClearTemporaryFiles()
    {
        _store.ClearTemporaryFiles();
        SetActionStatus("Временные извлечённые файлы очищены.");
    }

    private void PersistState()
    {
        _store.SaveState(_state);
        RefreshStatusBar();
    }

    private IEnumerable<DocumentRecord> FilteredDocuments(string? pageKey = null)
    {
        pageKey ??= _currentPage;
        var search = SearchTextBox.Text.Trim();
        var companyId = CurrentCompanyFilterId();
        var status = StatusFilterComboBox.IsEnabled ? SelectedString(StatusFilterComboBox) : null;
        var sectionId = SectionFilterComboBox.IsEnabled ? SelectedNullableInt(SectionFilterComboBox) : null;
        var problem = ProblemFilterComboBox.IsEnabled ? (SelectedString(ProblemFilterComboBox) ?? "all") : "all";

        var filtered = new List<DocumentRecord>();
        foreach (var document in _state.Documents)
        {
            var documentStatus = string.IsNullOrWhiteSpace(document.Status) ? AppConstants.DefaultStatus : document.Status;
            if (pageKey is "dashboard" or "attention" && string.Equals(documentStatus, "Архив", StringComparison.Ordinal))
            {
                continue;
            }

            if (pageKey == "documents" &&
                string.Equals(documentStatus, "Архив", StringComparison.Ordinal) &&
                !string.Equals(status, "Архив", StringComparison.Ordinal))
            {
                continue;
            }

            if (pageKey == "attention" && CalculateAttentionSeverity(document) <= 0 && !IsRecentUpdate(document.UpdatedAt))
            {
                continue;
            }

            if (companyId is int scopedCompanyId && document.CompanyId != scopedCompanyId)
            {
                continue;
            }

            if (!string.IsNullOrWhiteSpace(status) && !string.Equals(documentStatus, status, StringComparison.Ordinal))
            {
                continue;
            }

            if (sectionId is int scopedSectionId && !document.SectionIds.Contains(scopedSectionId))
            {
                continue;
            }

            if (!string.Equals(problem, "all", StringComparison.Ordinal) && !DocumentNeedsAttention(document, problem))
            {
                continue;
            }

            if (!string.IsNullOrWhiteSpace(search) && !MatchesDocumentSearch(document, search))
            {
                continue;
            }

            filtered.Add(document);
        }

        if (pageKey == "attention")
        {
            return filtered
                .OrderByDescending(document => CalculateAttentionSeverity(document))
                .ThenBy(document => IsRecentUpdate(document.UpdatedAt) ? 0 : 1)
                .ThenBy(document => ReviewPrioritySort(document))
                .ThenBy(document => document.ReviewDue ?? string.Empty, StringComparer.Ordinal)
                .ThenBy(document => GetCompanyName(document.CompanyId), StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(document => document.Title, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        return filtered
            .OrderBy(document => GetCompanyName(document.CompanyId), StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(document => document.SortOrder)
            .ThenBy(document => document.Title, StringComparer.CurrentCultureIgnoreCase)
            .ThenBy(document => document.Id)
            .ToList();
    }

    private bool MatchesDocumentSearch(DocumentRecord document, string search)
    {
        return ContainsIgnoreCase(document.Title, search) ||
               ContainsIgnoreCase(document.Comment, search) ||
               ContainsIgnoreCase(GetCompanyName(document.CompanyId), search) ||
               ContainsIgnoreCase(document.Status, search) ||
               SectionNames(document.SectionIds).Any(sectionName => ContainsIgnoreCase(sectionName, search));
    }

    private bool DocumentNeedsAttention(DocumentRecord document, string problemKey)
    {
        return problemKey switch
        {
            "missing-pdf" => !_store.RelativePathExists(document.PdfPath),
            "missing-office" => false,
            "due-review" => ReviewPriority(document) == "high",
            "upcoming-review" => ReviewPriority(document) is "mid" or "low",
            "missing-review" => TryParseStoredDate(document.ReviewDue) is null,
            "missing-section" => document.SectionIds.Count == 0,
            "recent-update" => IsRecentUpdate(document.UpdatedAt),
            _ => false,
        };
    }

    private IEnumerable<string> BuildIssueLines(DocumentRecord document)
    {
        if (DocumentNeedsAttention(document, "missing-pdf"))
        {
            yield return string.IsNullOrWhiteSpace(document.PdfPath) ? "PDF не загружен" : "PDF не найден";
        }

        if (DocumentNeedsAttention(document, "missing-section"))
        {
            yield return "документ без раздела";
        }

        switch (ReviewPriority(document))
        {
            case "high":
                yield return IsReviewDueToday(document) ? "пересмотр сегодня" : "просрочен пересмотр";
                break;
            case "mid":
                yield return "пересмотр в ближайшие 7 дней";
                break;
            case "low":
                yield return "пересмотр в ближайшие 30 дней";
                break;
            default:
                if (DocumentNeedsAttention(document, "missing-review"))
                {
                    yield return "не указана дата пересмотра";
                }
                break;
        }
    }

    private IEnumerable<string> BuildAttentionLines(DocumentRecord document)
    {
        if (DocumentNeedsAttention(document, "missing-pdf"))
        {
            yield return string.IsNullOrWhiteSpace(document.PdfPath) ? "PDF не загружен" : "PDF не найден";
        }

        if (DocumentNeedsAttention(document, "missing-section"))
        {
            yield return "Документ без раздела";
        }

        var dueText = FormatDisplayDate(document.ReviewDue, "не указана");
        switch (ReviewPriority(document))
        {
            case "high":
                yield return IsReviewDueToday(document)
                    ? $"Пересмотр сегодня: {dueText}"
                    : $"Просрочен пересмотр: {dueText}";
                break;
            case "mid":
                yield return $"Пересмотр в ближайшие 7 дней: {dueText}";
                break;
            case "low":
                yield return $"Пересмотр в ближайшие 30 дней: {dueText}";
                break;
            default:
                if (DocumentNeedsAttention(document, "missing-review"))
                {
                    yield return "Не указана дата пересмотра";
                }
                break;
        }
    }

    private int CalculateAttentionSeverity(DocumentRecord document)
    {
        var severity = 0;

        if (DocumentNeedsAttention(document, "missing-pdf"))
        {
            severity = Math.Max(severity, 5);
        }

        if (DocumentNeedsAttention(document, "missing-section"))
        {
            severity = Math.Max(severity, 2);
        }

        severity = ReviewPriority(document) switch
        {
            "high" => Math.Max(severity, 5),
            "mid" => Math.Max(severity, 3),
            "low" => Math.Max(severity, 2),
            _ when DocumentNeedsAttention(document, "missing-review") => Math.Max(severity, 1),
            _ => severity,
        };

        return severity;
    }

    private static int? ReviewDeltaDays(DocumentRecord document)
    {
        var due = TryParseStoredDate(document.ReviewDue);
        return due is null ? null : (due.Value.Date - DateTime.Today).Days;
    }

    private static bool IsReviewDueToday(DocumentRecord document)
    {
        return ReviewDeltaDays(document) == 0;
    }

    private string? ReviewPriority(DocumentRecord document)
    {
        var deltaDays = ReviewDeltaDays(document);
        if (deltaDays is null)
        {
            return null;
        }

        if (deltaDays.Value <= 0)
        {
            return "high";
        }

        if (deltaDays.Value <= 7)
        {
            return "mid";
        }

        if (deltaDays.Value <= 30)
        {
            return "low";
        }

        return null;
    }

    private int ReviewPrioritySort(DocumentRecord document)
    {
        return ReviewPriority(document) switch
        {
            "high" => 0,
            "mid" => 1,
            "low" => 2,
            _ => 3,
        };
    }

    private bool IsRecentUpdate(string? updatedAt, int days = 7)
    {
        var parsed = ParseSortTimestamp(updatedAt);
        return parsed != DateTimeOffset.MinValue && parsed >= DateTimeOffset.Now.AddDays(-days);
    }

    private string BuildScopeText(string pageKey)
    {
        var companyLabel = SelectedLabel(CompanyFilterComboBox) ?? "Все компании";
        var sectionLabel = SectionFilterComboBox.IsEnabled ? (SelectedLabel(SectionFilterComboBox) ?? "Все разделы") : null;
        var search = SearchTextBox.Text.Trim();

        var parts = new List<string> { companyLabel };
        if (!string.IsNullOrWhiteSpace(sectionLabel))
        {
            parts.Add(sectionLabel!);
        }
        if (!string.IsNullOrWhiteSpace(search))
        {
            parts.Add($"поиск: {search}");
        }

        return pageKey == "dashboard"
            ? $"Сейчас показываем: {string.Join(" · ", parts)}"
            : string.Join(" · ", parts);
    }

    private static void PopulateStackPanel(StackPanel panel, IEnumerable<UIElement> items, string emptyText)
    {
        panel.Children.Clear();
        var count = 0;
        foreach (var item in items)
        {
            panel.Children.Add(item);
            count++;
        }

        if (count == 0)
        {
            panel.Children.Add(new Border
            {
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0FFFFFFF")),
                BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#24C8AEFF")),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(16),
                Padding = new Thickness(14),
                Child = new TextBlock
                {
                    Text = emptyText,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B8B0C9")),
                    TextWrapping = TextWrapping.Wrap,
                },
            });
        }
    }

    private string BuildRecentDocumentCaption(DocumentRecord document)
    {
        var issue = BuildIssueLines(document).FirstOrDefault();
        if (!string.IsNullOrWhiteSpace(issue))
        {
            return issue;
        }

        return $"{GetCompanyName(document.CompanyId)} · {document.Status}";
    }

    private void RefreshCurrentPageForFilters()
    {
        switch (_currentPage)
        {
            case "dashboard":
                RefreshDashboardPage();
                break;
            case "attention":
                RefreshAttentionPage();
                break;
            case "documents":
                RefreshDocumentsPage();
                break;
            case "sopd":
                RefreshSopdPage();
                break;
        }
    }

    private void RefreshStatusBar()
    {
        ActionStatusTextBlock.Text ??= "Приложение готово к работе.";
    }

    private void SetActionStatus(string text)
    {
        ActionStatusTextBlock.Text = text;
    }

    private void OpenPath(string path)
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = path,
            UseShellExecute = true,
        });
    }

    private void DownloadAbsoluteFile(string? sourcePath, string dialogTitle, string successText)
    {
        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
        {
            ShowWarning("Файл не найден.");
            return;
        }

        var dialog = new SaveFileDialog
        {
            Title = dialogTitle,
            FileName = Path.GetFileName(sourcePath),
            Filter = "Все файлы (*.*)|*.*",
        };
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            File.Copy(sourcePath, dialog.FileName, true);
            SetActionStatus($"{successText} {dialog.FileName}");
        }
        catch (Exception ex)
        {
            ShowWarning($"Не удалось сохранить файл: {ex.Message}");
        }
    }

    private DocumentRecord? CurrentDocument()
    {
        return _selectedDocumentId is int documentId
            ? _state.Documents.FirstOrDefault(item => item.Id == documentId)
            : null;
    }

    private void ApplyDocumentFilters(string pageKey, int? companyId = null, string? problem = "all", string? status = null, int? sectionId = null, string? search = null)
    {
        _isRefreshingFilters = true;
        try
        {
            SearchTextBox.Text = search ?? string.Empty;
            SelectComboValue(CompanyFilterComboBox, companyId);
            SelectComboValue(StatusFilterComboBox, status);
            PopulateSectionFilterCombo(sectionId);
            SelectComboValue(ProblemFilterComboBox, problem ?? "all");
        }
        finally
        {
            _isRefreshingFilters = false;
        }

        SetPage(pageKey);
        if (_currentPage == pageKey)
        {
            RefreshCurrentPageForFilters();
        }
    }

    private void OpenCompanyScope(int companyId, string pageKey)
    {
        ApplyDocumentFilters(pageKey, companyId: companyId, problem: "all");
    }

    private int? CurrentCompanyFilterId() => SelectedNullableInt(CompanyFilterComboBox);

    private string GetCompanyName(int companyId)
    {
        return _state.Companies.FirstOrDefault(company => company.Id == companyId)?.Name ?? $"Компания #{companyId}";
    }

    private IEnumerable<string> SectionNames(IEnumerable<int> sectionIds)
    {
        var sectionMap = _state.Sections.ToDictionary(section => section.Id, section => section.Name);
        foreach (var sectionId in sectionIds)
        {
            if (sectionMap.TryGetValue(sectionId, out var name))
            {
                yield return name;
            }
        }
    }

    private static bool ContainsIgnoreCase(string? source, string value)
    {
        return !string.IsNullOrWhiteSpace(source) &&
               source.Contains(value, StringComparison.CurrentCultureIgnoreCase);
    }

    private static string NowIso() => DateTimeOffset.Now.ToString("O");

    private static DateTime? TryParseStoredDate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return DateTime.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)
            ? parsed
            : null;
    }

    private static string? ParseEditorDate(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return DateTime.TryParseExact(value, "dd.MM.yyyy", CultureInfo.GetCultureInfo("ru-RU"), DateTimeStyles.None, out var parsed)
            ? parsed.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : null;
    }

    private static string FormatEditorDate(string? value)
    {
        return TryParseStoredDate(value)?.ToString("dd.MM.yyyy", CultureInfo.GetCultureInfo("ru-RU")) ?? string.Empty;
    }

    private static string FormatDisplayDate(string? value, string fallback)
    {
        return TryParseStoredDate(value)?.ToString("dd.MM.yyyy", CultureInfo.GetCultureInfo("ru-RU")) ?? fallback;
    }

    private static string FormatTimestamp(string? value)
    {
        return DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)
            ? parsed.ToLocalTime().ToString("dd.MM.yyyy HH:mm", CultureInfo.GetCultureInfo("ru-RU"))
            : string.Empty;
    }

    private static string Csv(string value)
    {
        var escaped = value.Replace("\"", "\"\"", StringComparison.Ordinal);
        return $"\"{escaped}\"";
    }

    private Brush Brush(string key)
    {
        return (Brush)FindResource(key);
    }

    private static string OfficeVariant(string? path)
    {
        var extension = Path.GetExtension(path ?? string.Empty).ToLowerInvariant();
        return extension is ".xls" or ".xlsx" ? "excel" : "doc";
    }

    private static string OfficeBadgeText(string? path)
    {
        var extension = Path.GetExtension(path ?? string.Empty).TrimStart('.').ToUpperInvariant();
        return string.IsNullOrWhiteSpace(extension) ? "DOCX" : extension;
    }

    private static string? SelectedString(ComboBox comboBox)
    {
        return comboBox.SelectedItem is OptionItem option && option.Value is string value
            ? value
            : null;
    }

    private static int? SelectedNullableInt(ComboBox comboBox)
    {
        return comboBox.SelectedItem is OptionItem option && option.Value is int value
            ? value
            : null;
    }

    private static string? SelectedLabel(ComboBox comboBox)
    {
        return comboBox.SelectedItem is OptionItem option
            ? option.Label
            : null;
    }

    private static void SelectComboValue(ComboBox comboBox, object? value)
    {
        for (var index = 0; index < comboBox.Items.Count; index++)
        {
            if (comboBox.Items[index] is not OptionItem option)
            {
                continue;
            }

            if (Equals(option.Value, value))
            {
                comboBox.SelectedIndex = index;
                return;
            }
        }

        comboBox.SelectedIndex = comboBox.Items.Count > 0 ? 0 : -1;
    }

    private void ShowWarning(string message)
    {
        MessageDialog.ShowMessage(this, "Внимание", message, badgeText: "Внимание", useDangerAccent: true);
    }

    private bool Confirm(string message)
    {
        return MessageDialog.ShowConfirm(this, "Подтверждение", message, confirmText: "Да", cancelText: "Нет", useDangerAccent: true);
    }

    private void DashboardMetricCard_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (sender is not FrameworkElement element || element.Tag is not string tag)
        {
            return;
        }

        switch (tag)
        {
            case "documents":
                ApplyDocumentFilters("documents", companyId: CurrentCompanyFilterId(), problem: "all", status: SelectedString(StatusFilterComboBox), sectionId: SelectedNullableInt(SectionFilterComboBox), search: SearchTextBox.Text.Trim());
                break;
            default:
                ApplyDocumentFilters("attention", companyId: CurrentCompanyFilterId(), problem: tag, status: SelectedString(StatusFilterComboBox), sectionId: SelectedNullableInt(SectionFilterComboBox), search: SearchTextBox.Text.Trim());
                break;
        }
    }

    private void AttentionSummaryCard_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (sender is not FrameworkElement element || element.Tag is not string tag)
        {
            return;
        }

        ApplyDocumentFilters("attention", companyId: CurrentCompanyFilterId(), problem: tag, status: SelectedString(StatusFilterComboBox), sectionId: SelectedNullableInt(SectionFilterComboBox), search: SearchTextBox.Text.Trim());
    }

    private void NavButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is string pageKey)
        {
            SetPage(pageKey);
        }
    }

    private void PrimaryActionButton_Click(object sender, RoutedEventArgs e)
    {
        switch (_currentPage)
        {
            case "documents":
            case "attention":
                StartNewDocument();
                break;
            case "sopd":
                StartNewSopdRecord();
                break;
            default:
                AddCompany();
                break;
        }
    }

    private void DocumentFilter_Changed(object sender, RoutedEventArgs e)
    {
        if (_isRefreshingFilters)
        {
            return;
        }

        if (ReferenceEquals(sender, CompanyFilterComboBox))
        {
            PopulateSectionFilterCombo(SelectedNullableInt(SectionFilterComboBox));
        }

        RefreshCurrentPageForFilters();
    }

    private void ResetFiltersButton_Click(object sender, RoutedEventArgs e)
    {
        _isRefreshingFilters = true;
        try
        {
            SearchTextBox.Text = string.Empty;
            CompanyFilterComboBox.SelectedIndex = CompanyFilterComboBox.Items.Count > 0 ? 0 : -1;
            StatusFilterComboBox.SelectedIndex = StatusFilterComboBox.Items.Count > 0 ? 0 : -1;
            ProblemFilterComboBox.SelectedIndex = ProblemFilterComboBox.Items.Count > 0 ? 0 : -1;
            PopulateSectionFilterCombo(null);
        }
        finally
        {
            _isRefreshingFilters = false;
        }

        RefreshCurrentPageForFilters();
    }

    private void EditorCompanyComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isLoadingEditor)
        {
            return;
        }

        PopulateSectionsPanel(SelectedNullableInt(EditorCompanyComboBox), []);
        MarkDocumentDirty();
    }

    private void SaveDocumentButton_Click(object sender, RoutedEventArgs e)
    {
        SaveDocument();
    }

    private void DeleteDocumentButton_Click(object sender, RoutedEventArgs e)
    {
        DeleteDocument();
    }

    private void PdfUploadButton_Click(object sender, RoutedEventArgs e)
    {
        AttachFile("pdf");
    }

    private void PdfOpenButton_Click(object sender, RoutedEventArgs e)
    {
        OpenCurrentFile("pdf");
    }

    private void PdfDownloadButton_Click(object sender, RoutedEventArgs e)
    {
        DownloadCurrentFile("pdf");
    }

    private void PdfDeleteButton_Click(object sender, RoutedEventArgs e)
    {
        DeleteCurrentFile("pdf");
    }

    private void OfficeUploadButton_Click(object sender, RoutedEventArgs e)
    {
        AttachFile("office");
    }

    private void OfficeOpenButton_Click(object sender, RoutedEventArgs e)
    {
        OpenCurrentFile("office");
    }

    private void OfficeDownloadButton_Click(object sender, RoutedEventArgs e)
    {
        DownloadCurrentFile("office");
    }

    private void OfficeDeleteButton_Click(object sender, RoutedEventArgs e)
    {
        DeleteCurrentFile("office");
    }

    private void OpenWorkspaceFolder_Click(object sender, RoutedEventArgs e)
    {
        _store.OpenWorkspaceInExplorer();
    }

    private void ChangeWorkspaceFolder_Click(object sender, RoutedEventArgs e)
    {
        if (!MaybeDiscardSopdChanges() || !MaybeDiscardDocumentChanges())
        {
            return;
        }

        var dialog = new OpenFolderDialog
        {
            Title = "Выберите рабочую папку",
            InitialDirectory = _store.CurrentWorkspacePath,
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        _store.ChangeWorkspace(dialog.FolderName);
        _state = _store.LoadState();
        _selectedDocumentId = null;
        _selectedSopdId = null;
        _pendingPdfSourcePath = null;
        _pendingOfficeSourcePath = null;
        _pendingSopdAttachmentSourcePath = null;
        _isDocumentDirty = false;
        _isSopdDirty = false;
        RefreshAll();
        SetPage("dashboard");
        SetActionStatus($"Рабочая папка изменена: {_store.CurrentWorkspacePath}");
    }

    private void AddCompany_Click(object sender, RoutedEventArgs e)
    {
        AddCompany();
    }

    private void StartNewDocument_Click(object sender, RoutedEventArgs e)
    {
        StartNewDocument();
    }

    private void DuplicateCurrentDocument_Click(object sender, RoutedEventArgs e)
    {
        DuplicateCurrentDocument();
    }

    private void DuplicateCurrentSopd_Click(object sender, RoutedEventArgs e)
    {
        DuplicateCurrentSopd();
    }

    private void ExportDocuments_Click(object sender, RoutedEventArgs e)
    {
        ExportDocumentsToCsv();
    }

    private void CreateBackup_Click(object sender, RoutedEventArgs e)
    {
        CreateBackup();
    }

    private void RestoreBackup_Click(object sender, RoutedEventArgs e)
    {
        RestoreBackup();
    }

    private void ClearTempFiles_Click(object sender, RoutedEventArgs e)
    {
        ClearTemporaryFiles();
    }

    private void SidebarToggleButton_Click(object sender, RoutedEventArgs e)
    {
        ToggleSidebar();
    }

    private void SaveSopdButton_Click(object sender, RoutedEventArgs e)
    {
        SaveSopd();
    }

    private void DeleteSopdButton_Click(object sender, RoutedEventArgs e)
    {
        DeleteSopd();
    }

    private void SopdFileUploadButton_Click(object sender, RoutedEventArgs e)
    {
        AttachSopdFile();
    }

    private void SopdFileOpenButton_Click(object sender, RoutedEventArgs e)
    {
        OpenSopdFile();
    }

    private void SopdFileDownloadButton_Click(object sender, RoutedEventArgs e)
    {
        DownloadSopdFile();
    }

    private void SopdFileDeleteButton_Click(object sender, RoutedEventArgs e)
    {
        DeleteSopdFile();
    }

    private void SettingsCompanyListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        RefreshSettingsSections();
    }

    private void SettingsSectionListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        RefreshSettingsSummary();
    }

    private void SettingsAddCompanyButton_Click(object sender, RoutedEventArgs e)
    {
        AddCompany();
    }

    private void SettingsRenameCompanyButton_Click(object sender, RoutedEventArgs e)
    {
        RenameSelectedCompany();
    }

    private void SettingsCopyCompanyButton_Click(object sender, RoutedEventArgs e)
    {
        CopySelectedCompany();
    }

    private void SettingsDeleteCompanyButton_Click(object sender, RoutedEventArgs e)
    {
        DeleteSelectedCompany();
    }

    private void SettingsAddSectionButton_Click(object sender, RoutedEventArgs e)
    {
        AddSectionForSelectedCompany();
    }

    private void SettingsRenameSectionButton_Click(object sender, RoutedEventArgs e)
    {
        RenameSelectedSection();
    }

    private void SettingsDeleteSectionButton_Click(object sender, RoutedEventArgs e)
    {
        DeleteSelectedSection();
    }

    private void CompanySopdUploadButton_Click(object sender, RoutedEventArgs e)
    {
        UploadCompanySopdFile();
    }

    private void CompanySopdOpenButton_Click(object sender, RoutedEventArgs e)
    {
        OpenCompanySopdFile();
    }

    private void CompanySopdDownloadButton_Click(object sender, RoutedEventArgs e)
    {
        DownloadCompanySopdFile();
    }

    private void CompanySopdDeleteButton_Click(object sender, RoutedEventArgs e)
    {
        DeleteCompanySopdFile();
    }

    private sealed class OptionItem
    {
        public OptionItem(string label, object? value)
        {
            Label = label;
            Value = value;
        }

        public string Label { get; }

        public object? Value { get; }

        public override string ToString() => Label;
    }
}
