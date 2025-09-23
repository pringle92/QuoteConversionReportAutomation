// Form1.cs
// Main application form, fully updated with all new features and fixes.

#region Using Directives
// System related namespaces
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

#region Project Specific Namespaces

using QuoteConversionReportAutomation.Configuration;
using QuoteConversionReportAutomation.Forms;
using QuoteConversionReportAutomation.Helpers;
using QuoteConversionReportAutomation.Interfaces;
using QuoteConversionReportAutomation.Managers;
using QuoteConversionReportAutomation.Models;
using QuoteConversionReportAutomation.Models.Status;
using QuoteConversionReportAutomation.Orchestrators;
using QuoteConversionReportAutomation.Orchestrators.Interfaces;
using QuoteConversionReportAutomation.Services;
using QuoteConversionReportAutomation.Services.Communication;
using QuoteConversionReportAutomation.Services.Excel;
using QuoteConversionReportAutomation.Services.Interfaces;
using QuoteConversionReportAutomation.Services.Logging;
using QuoteConversionReportAutomation.Theming;
#endregion
#endregion

namespace conversionTest
{
    /// <summary>
    /// Represents the main form of the Quote Conversion Report Automation (QCRA) application.
    /// It serves as the primary user interface for manual report generation and for monitoring
    /// the automated report processes.
    /// </summary>
    public partial class Form1 : Form, IAutoRunUIContext
    {
        #region Fields and Properties
        private readonly IConfiguration _configuration;
        private readonly IReportPathService _reportPathService;
        private readonly IManualReportOrchestrator _manualReportOrchestrator;
        private readonly IRetrospectiveAnalysisOrchestrator _retrospectiveAnalysisOrchestrator;
        private readonly IBatchRegenerationOrchestrator _batchRegenerator;
        private readonly EmailUtility _emailUtility;
        private readonly ReportProcessManager _processManager;
        private readonly NamedPipeCommunicator _pipeCommunicator;
        private readonly AutoRunManager _autoRunManager;
        private readonly ExcelCopyData _excelProcessor;
        private readonly EmailRecipientManager _emailRecipientManager;
        private readonly GreetingManager _greetingManager;
        private readonly IServiceProvider _serviceProvider;
        private readonly IStatusManagerService _statusManager;
        private readonly UIManager _uiManager;
        private string _appName;
        private string _appVersion;
        private bool _programmaticallyChangingDates = false;
        private int _currentAutoRunHour;
        private HelpForm? _helpFormInstance;
        private string? _lastGeneratedRawReportPath = null;
        private string? _lastGeneratedAnalysisFilePath = null;
        private static bool IsDebug =>
#if DEBUG
            true;
#else
            false;
#endif
        #endregion

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="Form1"/> class, injecting all required dependencies.
        /// </summary>
        public Form1(
            IConfiguration configuration, IReportPathService reportPathService, IManualReportOrchestrator manualReportOrchestrator,
            IRetrospectiveAnalysisOrchestrator retrospectiveAnalysisOrchestrator, IBatchRegenerationOrchestrator batchRegenerator,
            EmailUtility emailUtility, ReportProcessManager processManager, NamedPipeCommunicator pipeCommunicator,
            AutoRunManager autoRunManager, ExcelCopyData excelProcessor, EmailRecipientManager emailRecipientManager,
            GreetingManager greetingManager, IServiceProvider serviceProvider, IStatusManagerService statusManager)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _reportPathService = reportPathService ?? throw new ArgumentNullException(nameof(reportPathService));
            _manualReportOrchestrator = manualReportOrchestrator ?? throw new ArgumentNullException(nameof(manualReportOrchestrator));
            _retrospectiveAnalysisOrchestrator = retrospectiveAnalysisOrchestrator ?? throw new ArgumentNullException(nameof(retrospectiveAnalysisOrchestrator));
            _batchRegenerator = batchRegenerator ?? throw new ArgumentNullException(nameof(batchRegenerator));
            _emailUtility = emailUtility ?? throw new ArgumentNullException(nameof(emailUtility));
            _processManager = processManager ?? throw new ArgumentNullException(nameof(processManager));
            _pipeCommunicator = pipeCommunicator ?? throw new ArgumentNullException(nameof(pipeCommunicator));
            _autoRunManager = autoRunManager ?? throw new ArgumentNullException(nameof(autoRunManager));
            _excelProcessor = excelProcessor ?? throw new ArgumentNullException(nameof(excelProcessor));
            _emailRecipientManager = emailRecipientManager ?? throw new ArgumentNullException(nameof(emailRecipientManager));
            _greetingManager = greetingManager ?? throw new ArgumentNullException(nameof(greetingManager));
            _serviceProvider = serviceProvider ?? throw new ArgumentNullException(nameof(serviceProvider));
            _statusManager = statusManager ?? throw new ArgumentNullException(nameof(statusManager));
            _appName = _configuration.GetValue<string>(AppConfigKeys.ApplicationInfo.AppName, "QCRA")!;
            _appVersion = _configuration.GetValue<string>(AppConfigKeys.ApplicationInfo.AppVersion, "1.0.0")!;

            InitializeComponent();

            _uiManager = new UIManager(this, menuStrip1, mainStatusStrip, autoRunStatusLabel,
                darkModeToolStripMenuItem, createReportButton, processEmailButton, oneClickProcessButton,
                toggleAutoRunButton, viewReportButton, viewAnalysisButton, reportTypeComboBox, startDatePicker,
                endDatePicker, financialYearComboBox, financialYearLabel, sendToFemiOnlyCheckBox,
                skipEmailCheckBox, chkIncludeLeadTimeAnalysis, emailRecipientLabel, toolTip1);

            _statusManager.StatusChanged += OnStatusChanged;
            _currentAutoRunHour = _configuration.GetValue<int>(AppConfigKeys.AutoRunProcess.CheckHour, 8);
            _uiManager.SetAutoRunHour(_currentAutoRunHour);
        }
        #endregion

        #region Status Event Handler and Interface Implementation
        /// <summary>
        /// Handles the StatusChanged event from the IStatusManagerService.
        /// This is the only place in the application that updates the main status label.
        /// </summary>
        private void OnStatusChanged(object? sender, StatusPayload payload)
        {
            UIManager.SafeToolStripItemUpdate(statusLabel, () =>
            {
                statusLabel.Text = payload.Message;
                statusLabel.ForeColor = payload.Type switch
                {
                    MessageType.Success => Color.Green,
                    MessageType.Warning => Color.Goldenrod,
                    MessageType.Error => Color.Firebrick,
                    _ => ThemeSettings.CurrentPalette.StatusStripForeColor
                };
            });
        }

        /// <inheritdoc/>
        public void ReportAutoRunProgress(string message) => _statusManager.Post(message, MessageType.InProgress);

        /// <inheritdoc/>
        public void ReportAutoRunStatusRight(string message) => _uiManager.UpdateStatusRight(message);

        /// <inheritdoc/>
        public void SetControlsForAutoRunInProgress(bool inProgress) { if (inProgress) _uiManager.DisableControlsForAutoRun(); }

        /// <inheritdoc/>
        public void UpdateAutoRunButtonAndStatus(bool isTimerEnabled, bool isJobDoneOrFailedForToday, string statusTextToDisplay) => _uiManager.UpdateAutoRunUI(isTimerEnabled, isJobDoneOrFailedForToday, statusTextToDisplay);

        /// <inheritdoc/>
        public bool IsWindowsDarkModeEnabled() => ThemeSettings.IsWindowsDarkModeEnabled();
        #endregion

        #region Form Lifecycle and Main Action Handlers
        /// <summary>
        /// Handles the main form's Load event.
        /// </summary>
        private async void Form1_Load(object sender, EventArgs e)
        {
            _statusManager.Post("Loading application...", MessageType.InProgress);
            PopulateReportTypeComboBox();
            BankHolidayHelper.Initialize();
            bool configValid = _reportPathService.IsEssentialPathConfigurationValid();
            Text = $"{_appName} - {(IsDebug ? "DEBUG" : "RELEASE")} - v{_appVersion}";
            StartPosition = FormStartPosition.CenterScreen;
            financialYearComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            reportTypeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            reportTypeComboBox.SelectedItem = "Daily";
            ThemeSettings.SyncThemeWithSystem();
            darkModeToolStripMenuItem.Checked = ThemeSettings.IsCurrentlyDark();
            _uiManager.ApplyTheme();
            UpdateAutoRunButtonAndStatus(dailyCheckTimer.Enabled, false, $"Auto Run: {(dailyCheckTimer.Enabled ? $"Enabled (Next check ~{_currentAutoRunHour}:00)" : "Disabled")}");
            reportTypeComboBox_SelectedIndexChanged(reportTypeComboBox, EventArgs.Empty);
            _uiManager.ResetButtonStatesAfterTypeChange(configValid);
            Update1ClickProcessingModeUI();
            if (!configValid) _statusManager.Post("Config Error: Check Options menu.", MessageType.Error);
            _statusManager.Post("Checking report service...", MessageType.InProgress);
            bool wrapperOk = await _processManager.EnsureWrapperIsRunningAsync(new Progress<string>(status => _statusManager.Post(status, MessageType.InProgress)));
            if (!wrapperOk && configValid) _statusManager.Post("Report service failed to start. Report generation may fail.", MessageType.Warning);
            _ = Task.Run(() => ReportArchiver.ArchiveOldReportsAsync(_reportPathService.FinalReportOutputBaseDirectory, _reportPathService.RawReportExportBaseDirectory, _configuration.GetValue<int?>(AppConfigKeys.OperationalParameters.ArchiveRawReportsOlderThanDays), _configuration.GetValue<string>(AppConfigKeys.OperationalParameters.ReportArchiveFolderName)));
            if (configValid && wrapperOk) _statusManager.Clear();
            else if (configValid && !wrapperOk) _statusManager.Post("Ready (Report Service Issue)", MessageType.Warning);
            else _statusManager.Post("Config Error (Service Check Skipped)", MessageType.Error);
        }

        /// <summary>
        /// Handles the form's Closing event to shut down background processes.
        /// </summary>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            dailyCheckTimer.Stop();
            _processManager.TerminateWrapperProcess();
            _statusManager.StatusChanged -= OnStatusChanged;
        }

        /// <summary>
        /// Handles the Click event for the "Create Report" button.
        /// </summary>
        private async void createReportButton_Click(object sender, EventArgs e)
        {
            if (!ValidateInputDates() || !ValidateFinancialYearSelection()) { _uiManager.ResetUIOnError("Create Report", _reportPathService.IsEssentialPathConfigurationValid(), File.Exists(_lastGeneratedRawReportPath), File.Exists(_lastGeneratedAnalysisFilePath), IsAnyDailySelected(), dailyCheckTimer.Enabled, false, _uiManager.GetAutoRunStatusLabelText()); return; }
            _statusManager.Post("Requesting raw report...", MessageType.InProgress);
            _uiManager.SetActionButtonsEnabled(false);
            _uiManager.SetOtherControlsEnabled(false, financialYearComboBox.Visible);
            var parameters = GatherManualReportParameters();
            using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(_configuration.GetValue<int>(AppConfigKeys.OperationalParameters.ProcessTimeoutMinutes, 6)));
            ReportCreationResult result = await _manualReportOrchestrator.CreateRawReportAsync(parameters, cts.Token);
            if (result.Success && !string.IsNullOrEmpty(result.GeneratedRawPath))
            {
                _lastGeneratedRawReportPath = result.GeneratedRawPath;
                _lastGeneratedAnalysisFilePath = null;
                _uiManager.ShowViewReportButton(true, _lastGeneratedRawReportPath);
                _uiManager.ShowViewAnalysisButton(false);
                _statusManager.Post("Raw report created successfully.", MessageType.Success, TimeSpan.FromSeconds(5));
                UIManager.SafeControlUpdate(createReportButton, () => createReportButton.Text = "Report Created");
                UIManager.SafeControlUpdate(processEmailButton, () => processEmailButton.Enabled = true);
            }
            else
            {
                _lastGeneratedRawReportPath = null;
                FlexibleMessageBox.Show(this, result.ErrorMessage ?? "Raw report creation failed for an unknown reason.", "Report Creation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _statusManager.Post(result.ErrorMessage ?? "Raw report creation failed.", MessageType.Error);
            }
            if (!oneClickProcessButton.Visible || !result.Success)
            {
                _uiManager.ResetUIOnError("Create Report", _reportPathService.IsEssentialPathConfigurationValid(), File.Exists(_lastGeneratedRawReportPath), File.Exists(_lastGeneratedAnalysisFilePath), IsAnyDailySelected(), dailyCheckTimer.Enabled, false, _uiManager.GetAutoRunStatusLabelText());
            }
        }

        /// <summary>
        /// Handles the Click event for the "Process & Email" button.
        /// </summary>
        private async void processEmailButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_lastGeneratedRawReportPath) || !File.Exists(_lastGeneratedRawReportPath)) { FlexibleMessageBox.Show(this, "The raw report file has not been generated or cannot be found. Please create the report first.", "Raw Report Missing", MessageBoxButtons.OK, MessageBoxIcon.Warning); _uiManager.ResetUIOnError("Create Report", _reportPathService.IsEssentialPathConfigurationValid(), false, File.Exists(_lastGeneratedAnalysisFilePath), IsAnyDailySelected(), dailyCheckTimer.Enabled, false, _uiManager.GetAutoRunStatusLabelText()); return; }
            if (!ValidateInputDates() || !ValidateFinancialYearSelection()) { _uiManager.ResetUIOnError("Process & Email", _reportPathService.IsEssentialPathConfigurationValid(), File.Exists(_lastGeneratedRawReportPath), File.Exists(_lastGeneratedAnalysisFilePath), IsAnyDailySelected(), dailyCheckTimer.Enabled, false, _uiManager.GetAutoRunStatusLabelText()); return; }
            _statusManager.Post("Processing report and preparing email...", MessageType.InProgress);
            _uiManager.SetActionButtonsEnabled(false);
            _uiManager.SetOtherControlsEnabled(false, financialYearComboBox.Visible);
            var parameters = GatherManualReportParameters();
            using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(_configuration.GetValue<int>(AppConfigKeys.OperationalParameters.ProcessTimeoutMinutes, 15)));
            ReportProcessingResult result = await _manualReportOrchestrator.ProcessAndEmailReportAsync(_lastGeneratedRawReportPath, parameters, cts.Token);
            await HandleReportProcessingResult(result, parameters, cts.Token);
        }

        /// <summary>
        /// Handles the Click event for the "1-Click Process" button.
        /// </summary>
        private async void oneClickProcessButton_Click(object sender, EventArgs e)
        {
            if (!ValidateInputDates() || !ValidateFinancialYearSelection()) { _uiManager.ResetUIOnError("Generate, Process & Email Report", _reportPathService.IsEssentialPathConfigurationValid(), File.Exists(_lastGeneratedRawReportPath), File.Exists(_lastGeneratedAnalysisFilePath), IsAnyDailySelected(), dailyCheckTimer.Enabled, false, _uiManager.GetAutoRunStatusLabelText()); return; }
            _statusManager.Post("1-Click Process: Starting...", MessageType.InProgress);
            _uiManager.SetActionButtonsEnabled(false);
            _uiManager.SetOtherControlsEnabled(false, financialYearComboBox.Visible);
            var parameters = GatherManualReportParameters();
            using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(_configuration.GetValue<int>(AppConfigKeys.OperationalParameters.ProcessTimeoutMinutes, 15)));
            ReportCreationResult creationResult = await _manualReportOrchestrator.CreateRawReportAsync(parameters, cts.Token);
            if (!creationResult.Success || string.IsNullOrEmpty(creationResult.GeneratedRawPath))
            {
                _lastGeneratedRawReportPath = null;
                FlexibleMessageBox.Show(this, creationResult.ErrorMessage ?? "Raw report creation failed in 1-Click process.", "1-Click Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _uiManager.ResetUIOnError("Generate, Process & Email Report", _reportPathService.IsEssentialPathConfigurationValid(), false, File.Exists(_lastGeneratedAnalysisFilePath), IsAnyDailySelected(), dailyCheckTimer.Enabled, false, _uiManager.GetAutoRunStatusLabelText());
                return;
            }
            _lastGeneratedRawReportPath = creationResult.GeneratedRawPath;
            _uiManager.ShowViewReportButton(true, _lastGeneratedRawReportPath);
            _statusManager.Post("1-Click: Raw report created. Processing...", MessageType.InProgress);
            ReportProcessingResult processingResult = await _manualReportOrchestrator.ProcessAndEmailReportAsync(_lastGeneratedRawReportPath, parameters, cts.Token);
            await HandleReportProcessingResult(processingResult, parameters, cts.Token);
        }

        /// <summary>
        /// Handles the result of a report processing operation, updating UI and showing messages.
        /// </summary>
        private async Task HandleReportProcessingResult(ReportProcessingResult result, ManualReportParameters parameters, CancellationToken originalCts)
        {
            if (result.Success)
            {
                _lastGeneratedAnalysisFilePath = result.GeneratedAnalysisPath;
                _uiManager.ShowViewAnalysisButton(true, _lastGeneratedAnalysisFilePath);

                if (result.EmailResult?.Success == true || parameters.SkipEmail) { _statusManager.Post("Process completed successfully.", MessageType.Success, TimeSpan.FromSeconds(5)); }
                else if (result.EmailResult?.Success == false) { FlexibleMessageBox.Show(this, result.EmailResult.ErrorMessage ?? "Email sending failed.", "Email Error", MessageBoxButtons.OK, MessageBoxIcon.Error); _statusManager.Post(result.EmailResult.ErrorMessage ?? "Email sending failed.", MessageType.Error); }
                else { _statusManager.Post("Processing complete. Email status unknown.", MessageType.Warning, TimeSpan.FromSeconds(5)); }
            }
            else
            {
                _lastGeneratedAnalysisFilePath = null;
                FlexibleMessageBox.Show(this, result.ErrorMessage ?? "Report processing failed for an unknown reason.", "Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _statusManager.Post(result.ErrorMessage ?? "Report processing failed.", MessageType.Error);
            }
            _uiManager.ResetUIOnError(oneClickProcessButton.Visible ? "Generate, Process & Email Report" : "Create Report", _reportPathService.IsEssentialPathConfigurationValid(), File.Exists(_lastGeneratedRawReportPath), File.Exists(_lastGeneratedAnalysisFilePath), IsAnyDailySelected(), dailyCheckTimer.Enabled, false, _uiManager.GetAutoRunStatusLabelText());
        }

        /// <summary>
        /// Handles the Click event for the "View Raw Report" button.
        /// </summary>
        private void viewReportButton_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_lastGeneratedRawReportPath))
            {
                try { ReportHelper.OpenFileWithDefaultApp(_lastGeneratedRawReportPath, "raw report output"); }
                catch (Exception ex) { FlexibleMessageBox.Show(this, ex.Message, "Error Opening File", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            else
            {
                FlexibleMessageBox.Show(this, "No raw report has been generated yet in this session.", "File Not Available", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Handles the Click event for the "View Processed Analysis" button.
        /// </summary>
        private void viewAnalysisButton_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_lastGeneratedAnalysisFilePath))
            {
                try { ReportHelper.OpenFileWithDefaultApp(_lastGeneratedAnalysisFilePath, "processed analysis file"); }
                catch (Exception ex) { FlexibleMessageBox.Show(this, ex.Message, "Error Opening File", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            else
            {
                FlexibleMessageBox.Show(this, "No analysis file has been generated or successfully processed yet in this session.", "File Not Available", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        #endregion

        #region UI Event Handlers (ComboBoxes, DatePickers, MenuItems)
        /// <summary>
        /// Populates the report type combo box with dynamically generated display names.
        /// </summary>
        private void PopulateReportTypeComboBox()
        {
            reportTypeComboBox.Items.Clear();
            foreach (ReportType type in Enum.GetValues(typeof(ReportType)))
            {
                if (type == ReportType.Unknown) continue;
                reportTypeComboBox.Items.Add(ReportTypeHelper.GetDisplayString(type, _configuration));
            }
        }

        /// <summary>
        /// Handles the SelectedIndexChanged event for the report type ComboBox.
        /// </summary>
        private void reportTypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (sender is not ComboBox comboBox || comboBox.SelectedItem == null) return;

            ReportType selectedReportType = GetSelectedReportType();

            if (selectedReportType == ReportType.Custom)
            {
                UIManager.SafeControlUpdate(sendToFemiOnlyCheckBox, () => { sendToFemiOnlyCheckBox.Visible = true; });
                UIManager.SafeControlUpdate(emailRecipientLabel, () => { emailRecipientLabel.Visible = false; });
                _uiManager.ResetButtonStatesAfterTypeChange(_reportPathService.IsEssentialPathConfigurationValid());
                Update1ClickProcessingModeUI();
                return;
            }

            DateTime todayValue = DateTime.Today;
            _programmaticallyChangingDates = true;
            try
            {
                (DateTime dateFrom, DateTime dateTo, bool showFinYear) = selectedReportType switch
                {
                    ReportType.Daily => (ReportHelper.GetPreviousWorkday(todayValue), ReportHelper.GetPreviousWorkday(todayValue), false),
                    ReportType.Daily5Day1k => (ReportHelper.GetNthPreviousWorkday(ReportHelper.GetPreviousWorkday(todayValue), 4), ReportHelper.GetPreviousWorkday(todayValue), false),
                    ReportType.Weekly => (todayValue.AddDays(-14), todayValue, true),
                    ReportType.Monthly => (ReportHelper.CalculateMonthlyRange(todayValue).DateFrom, ReportHelper.CalculateMonthlyRange(todayValue).DateTo, false),
                    ReportType.Quarterly => (ReportHelper.CalculateQuarterlyRange(todayValue).DateFrom, ReportHelper.CalculateQuarterlyRange(todayValue).DateTo, false),
                    ReportType.Annual => (ReportHelper.GetFinancialYearDates(ReportHelper.GetFinancialYearStartCalendarYear(todayValue, _configuration) - 1, _configuration).DateFrom, ReportHelper.GetFinancialYearDates(ReportHelper.GetFinancialYearStartCalendarYear(todayValue, _configuration) - 1, _configuration).DateTo, false),
                    _ => (startDatePicker.Value, endDatePicker.Value, true)
                };

                UIManager.SafeControlUpdate(startDatePicker, () => { startDatePicker.Value = dateFrom; });
                UIManager.SafeControlUpdate(endDatePicker, () => { endDatePicker.Value = dateTo; });
                UIManager.SafeControlUpdate(financialYearLabel, () => { financialYearLabel.Visible = showFinYear; });
                UIManager.SafeControlUpdate(financialYearComboBox, () => { financialYearComboBox.Visible = showFinYear; financialYearComboBox.Enabled = showFinYear; if (showFinYear) PopulateFinancialYearDropdown(); });

                // Updated visibility logic for recipient controls
                bool isStandardDailyOnly = selectedReportType == ReportType.Daily;
                UIManager.SafeControlUpdate(emailRecipientLabel, () =>
                {
                    emailRecipientLabel.Visible = isStandardDailyOnly;
                    if (isStandardDailyOnly)
                    {
                        emailRecipientLabel.Text = "Manual Daily: Uses configured list.";
                    }
                });

                UIManager.SafeControlUpdate(sendToFemiOnlyCheckBox, () => { sendToFemiOnlyCheckBox.Visible = (selectedReportType != ReportType.Daily && selectedReportType != ReportType.Custom); });

                _uiManager.ResetButtonStatesAfterTypeChange(_reportPathService.IsEssentialPathConfigurationValid());
                Update1ClickProcessingModeUI();
            }
            finally
            {
                _programmaticallyChangingDates = false;
            }
        }

        /// <summary>
        /// Handles date changes to switch the report type to "Custom" if appropriate.
        /// This now intelligently avoids changing the type if it's already a type
        /// that implies manual date entry (like Custom or NewCustomer).
        /// </summary>
        private void DatePicker_ValueChanged(object sender, EventArgs e)
        {
            if (_programmaticallyChangingDates) return;

            ReportType currentType = GetSelectedReportType();

            // Only switch to 'Custom' if the current type is a pre-defined, fixed-date range type.
            // Leave 'Custom' and 'NewCustomer' as they are, since they imply manual date entry.
            if (currentType is ReportType.Daily or ReportType.Daily5Day1k or ReportType.Weekly or ReportType.Monthly or ReportType.Quarterly or ReportType.Annual)
            {
                UIManager.SafeControlUpdate(reportTypeComboBox, () =>
                {
                    reportTypeComboBox.SelectedItem = ReportTypeHelper.GetDisplayString(ReportType.Custom, _configuration);
                });
            }
        }

        /// <summary>
        /// Handles the Click event for the Dark Mode menu item.
        /// </summary>
        private void darkModeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ThemeSettings.CurrentThemeMode = darkModeToolStripMenuItem.Checked ? ApplicationThemeMode.Dark : ApplicationThemeMode.Light;
            _uiManager.ApplyTheme();
        }

        /// <summary>
        /// Handles the Click event for the main "Settings..." menu item.
        /// </summary>
        private void settingsToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            using var settingsFormInstance = _serviceProvider.GetRequiredService<SettingsForm>();
            if (settingsFormInstance.ShowDialog(this) == DialogResult.OK)
            {
                if (_configuration is IConfigurationRoot configurationRoot)
                {
                    configurationRoot.Reload();
                    FlexibleMessageBox.Show(this, "Settings saved and configuration reloaded.\nA restart may be needed for some changes to fully apply.", "Settings Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ReinitializeConfigurableComponents();
                }
                else
                {
                    FlexibleMessageBox.Show(this, "Settings saved. Please restart the application for changes to take effect.", "Settings Saved - Restart Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        /// <summary>
        /// Handles the Click event for the "Help" menu item.
        /// </summary>
        private void helpToolStripMenuItem_Click(object? sender, EventArgs e)
        {
            string helpTitle = ReportHelper.GetHelpTitle(_appName, _appVersion);
            string helpContent = ReportHelper.GetHelpContent(_configuration, _appName, _appVersion);
            try
            {
                if (_helpFormInstance == null || _helpFormInstance.IsDisposed)
                {
                    _helpFormInstance = new HelpForm(helpTitle, helpContent);
                    _helpFormInstance.FormClosed += (s, args) => _helpFormInstance = null;
                    _helpFormInstance.Show(this);
                }
                else
                {
                    _helpFormInstance.Activate();
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"Failed to show HelpForm: {ex.Message}", ex);
                FlexibleMessageBox.Show(this, "Could not display help window. Please check application logs.", "Help Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Handles the Click event for the "View Configuration" menu item.
        /// </summary>
        private void viewConfigToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool configValid = _reportPathService.IsEssentialPathConfigurationValid();
            var sb = new StringBuilder();
            sb.AppendLine("Configuration Details (Paths are relative to user profile where applicable):").AppendLine("--------------------------------------------------").AppendLine($"1. Crystal Report Path (.rpt): '{_reportPathService.CrystalReportRptFilePath}' - Exists: {File.Exists(_reportPathService.CrystalReportRptFilePath)}").AppendLine($"2. Wrapper EXE Path: '{_reportPathService.WrapperExecutablePath}' - Exists: {File.Exists(_reportPathService.WrapperExecutablePath)}").AppendLine($"3. Template Base Directory: '{_reportPathService.TemplateBaseDirectory}' - Exists: {Directory.Exists(_reportPathService.TemplateBaseDirectory)}").AppendLine($"4. Raw Report Export Base Directory: '{_reportPathService.RawReportExportBaseDirectory}' - Exists: {Directory.Exists(_reportPathService.RawReportExportBaseDirectory)}").AppendLine($"5. Final Excel Save Location Base: '{_reportPathService.FinalReportOutputBaseDirectory}' - Exists: {Directory.Exists(_reportPathService.FinalReportOutputBaseDirectory)}").AppendLine($"6. Auto-Run Check Hour: {_configuration.GetValue<int>(AppConfigKeys.AutoRunProcess.CheckHour, _currentAutoRunHour)} (Current in-memory: {_currentAutoRunHour})").AppendLine($"7. Automated Report Definitions File: '{_reportPathService.GetReportDefinitionsFilePath() ?? "N/A"}' - Exists: {File.Exists(_reportPathService.GetReportDefinitionsFilePath() ?? string.Empty)}").AppendLine($"8. Application Log Directory (User Specific): '{_reportPathService.GetUserSpecificLogDirectory()}' - Exists: {Directory.Exists(_reportPathService.GetUserSpecificLogDirectory())}").AppendLine($"9. appsettings.json Directory: '{_reportPathService.AppSettingsDirectory}' - appsettings.json Exists: {File.Exists(Path.Combine(_reportPathService.AppSettingsDirectory, "appsettings.json"))}").AppendLine("--------------------------------------------------").AppendLine($"Overall Essential Config Valid (for report generation): {configValid}");
            FlexibleMessageBox.Show(this, sb.ToString(), "Configuration Details", MessageBoxButtons.OK, configValid ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
        }

        /// <summary>
        /// Handles the Click event for the "Validate Configuration" menu item.
        /// </summary>
        private void validateConfigToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _statusManager.Post("Validating configuration...", MessageType.InProgress);
            bool isValid = _reportPathService.IsEssentialPathConfigurationValid();
            string statusMessage = isValid ? "Configuration OK." : "Configuration Error: Essential paths missing or invalid.";
            MessageType type = isValid ? MessageType.Success : MessageType.Error;
            _statusManager.Post(statusMessage, type, TimeSpan.FromSeconds(5));
            if (!isValid) Logger.LogError("Configuration validation failed.");
        }

        /// <summary>
        /// Handles the Click event for the "Open Logs" menu item.
        /// </summary>
        private void openLogsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string userLogDir = _reportPathService.GetUserSpecificLogDirectory();
                if (!Directory.Exists(userLogDir)) Directory.CreateDirectory(userLogDir);
                Process.Start("explorer.exe", userLogDir);
            }
            catch (Exception ex)
            {
                FlexibleMessageBox.Show(this, $"Could not open logs folder: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Handles the Click event for the "Edit Config" menu item.
        /// </summary>
        private void editConfigToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string appSettingsJsonPath = Path.Combine(_reportPathService.AppSettingsDirectory, "appsettings.json");
                if (File.Exists(appSettingsJsonPath)) Process.Start(new ProcessStartInfo(appSettingsJsonPath) { UseShellExecute = true });
                else FlexibleMessageBox.Show(this, $"appsettings.json not found at '{appSettingsJsonPath}'", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                FlexibleMessageBox.Show(this, $"Could not open appsettings.json: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Handles the Click event for the "Exit" menu item.
        /// </summary>
        private void exitToolStripMenuItem_Click(object sender, EventArgs e) => Close();

        /// <summary>
        /// Handles the Click event for the "Manage Bank Holidays" menu item.
        /// </summary>
        private void manageCustomBankHolidaysToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                using var form = _serviceProvider.GetRequiredService<ManageBankHolidaysForm>();
                form.ShowDialog(this);
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error opening ManageBankHolidaysForm: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Handles the Click event for the "Manage Email Recipients" menu item.
        /// </summary>
        private void manageEmailRecipientsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                using var form = _serviceProvider.GetRequiredService<ManageEmailRecipientsForm>();
                form.ShowDialog(this);
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error opening ManageEmailRecipientsForm: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Handles the Click event for the "Manage Greetings" menu item.
        /// </summary>
        private void manageGreetingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                using var form = _serviceProvider.GetRequiredService<ManageGreetingsForm>();
                form.ShowDialog(this);
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error opening ManageGreetingsForm: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Handles the Click event for the "1-Click Processing" menu item.
        /// </summary>
        private void enable1ClickProcessingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Update1ClickProcessingModeUI();
            string mainButtonTextForReset = enable1ClickProcessingToolStripMenuItem.Checked ? (_reportPathService.IsEssentialPathConfigurationValid() ? "Generate, Process & Email Report" : "Config Error") : (_reportPathService.IsEssentialPathConfigurationValid() ? "Create Report" : "Config Error");
            _uiManager.ResetUIOnError(mainButtonTextForReset, _reportPathService.IsEssentialPathConfigurationValid(), File.Exists(_lastGeneratedRawReportPath), File.Exists(_lastGeneratedAnalysisFilePath), IsAnyDailySelected(), dailyCheckTimer.Enabled, false, _uiManager.GetAutoRunStatusLabelText());
        }

        /// <summary>
        /// Handles the Click event for the "Set Auto-Run Hour" menu item.
        /// </summary>
        private async void setAutoRunHourToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Use the Microsoft.VisualBasic.Interaction.InputBox for simple text input.
            string? inputText = Microsoft.VisualBasic.Interaction.InputBox("Enter new hour (0-23) for daily auto-run check:", "Set Auto-Run Hour", _currentAutoRunHour.ToString());

            if (int.TryParse(inputText, out int newHour) && newHour >= 0 && newHour <= 23)
            {
                if (newHour != _currentAutoRunHour && await _autoRunManager.SetAutoRunHourAsync(newHour))
                {
                    _currentAutoRunHour = newHour;
                    _uiManager.SetAutoRunHour(_currentAutoRunHour);
                    FlexibleMessageBox.Show(this, $"Auto-Run hour set to {newHour}:00.", "Auto-Run Hour Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    _uiManager.UpdateAutoRunUI(dailyCheckTimer.Enabled, false, $"Auto Run: {(dailyCheckTimer.Enabled ? $"Enabled (Next check ~{_currentAutoRunHour}:00)" : "Disabled")}");
                }
            }
            else if (!string.IsNullOrWhiteSpace(inputText)) // Show error only if user entered something invalid
            {
                FlexibleMessageBox.Show(this, "Invalid hour. Please enter a number between 0 and 23.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Handles the Click event for the "Manage Automated Reports" menu item.
        /// </summary>
        private void manageAutomatedReportsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                using var form = _serviceProvider.GetRequiredService<ManageAutoReportDefinitionsForm>();
                form.ShowDialog(this);
                _autoRunManager.ReloadReportDefinitions();
                _autoRunManager.SynchronizeSuccessFlags();
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error opening ManageAutoReportDefinitionsForm: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Handles the Click event for the "Open Auto Report Definitions" menu item.
        /// </summary>
        private void openAutoReportDefinitionsFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string? filePath = _reportPathService.GetReportDefinitionsFilePath();
                if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath)) Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                else FlexibleMessageBox.Show(this, $"Auto report definitions file not found at: {filePath ?? "N/A"}", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error opening auto report definitions file: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Handles the Click event for the "Retrospective Analysis" menu item.
        /// </summary>
        private async void retrospectiveAnalysisToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var optionsForm = _serviceProvider.GetRequiredService<AnalysisOptionsForm>();
            if (optionsForm.ShowDialog(this) == DialogResult.OK)
            {
                _uiManager.SetActionButtonsEnabled(false);
                _uiManager.SetOtherControlsEnabled(false, financialYearComboBox.Visible);

                try
                {
                    await _retrospectiveAnalysisOrchestrator.GenerateAnalysisAsync(
                        optionsForm.SelectedFolder,
                        optionsForm.FileNamePattern,
                        CancellationToken.None);
                }
                catch (Exception ex)
                {
                    _statusManager.Post($"Analysis failed: {ex.Message}", MessageType.Error);
                    Logger.LogError("Retrospective Analysis failed with a critical exception.", ex);
                }
                finally
                {
                    _uiManager.ResetUIOnError("Create Report", _reportPathService.IsEssentialPathConfigurationValid(), File.Exists(_lastGeneratedRawReportPath), File.Exists(_lastGeneratedAnalysisFilePath), IsAnyDailySelected(), dailyCheckTimer.Enabled, false, _uiManager.GetAutoRunStatusLabelText());
                }
            }
        }

        /// <summary>
        /// Handles the Click event for the "Batch Regenerate Reports" menu item.
        /// </summary>
        private async void batchRegenerateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var dateForm = _serviceProvider.GetRequiredService<DateRangeSelectionForm>();
            if (dateForm.ShowDialog(this) == DialogResult.OK)
            {
                ReportType selectedType = dateForm.SelectedReportType;
                if (selectedType == ReportType.Unknown)
                {
                    FlexibleMessageBox.Show(this, "You must select a valid report type.", "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (dateForm.EndDate < dateForm.StartDate)
                {
                    FlexibleMessageBox.Show(this, "The end date cannot be before the start date.", "Invalid Date Range", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var confirmResult = FlexibleMessageBox.Show(this,
                    $"This will regenerate all '{selectedType}' reports from {dateForm.StartDate:d} to {dateForm.EndDate:d}.\n\nThis can take a very long time and will overwrite existing files. Are you sure you want to continue?",
                    "Confirm Batch Regeneration", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (confirmResult == DialogResult.No) return;

                _uiManager.SetActionButtonsEnabled(false);
                _uiManager.SetOtherControlsEnabled(false, financialYearComboBox.Visible);

                try
                {
                    await _batchRegenerator.RegenerateReportsAsync(selectedType, dateForm.StartDate, dateForm.EndDate, CancellationToken.None);
                }
                catch (Exception ex)
                {
                    _statusManager.Post($"Batch regeneration failed: {ex.Message}", MessageType.Error);
                    Logger.LogError("Batch regeneration failed with a critical exception.", ex);
                }
                finally
                {
                    _uiManager.ResetUIOnError("Create Report", _reportPathService.IsEssentialPathConfigurationValid(), File.Exists(_lastGeneratedRawReportPath), File.Exists(_lastGeneratedAnalysisFilePath), IsAnyDailySelected(), dailyCheckTimer.Enabled, false, _uiManager.GetAutoRunStatusLabelText());
                }
            }
        }
        #endregion

        #region Auto-Run Timer Event Handler
        /// <summary>
        /// Handles the Tick event for the daily auto-run timer.
        /// </summary>
        private async void dailyCheckTimer_Tick(object sender, EventArgs e)
        {
            if (!dailyCheckTimer.Enabled || _autoRunManager == null) return;
            bool originallyEnabled = dailyCheckTimer.Enabled;
            dailyCheckTimer.Stop();
            AutoRunActionResult autoRunResult = AutoRunActionResult.NoActionNeeded;
            try
            {
                autoRunResult = await _autoRunManager.PerformDailyCheckAsync(originallyEnabled, _currentAutoRunHour);
            }
            catch (Exception ex)
            {
                Logger.LogCritical($"CRITICAL ERROR during AutoRunManager.PerformDailyCheckAsync dispatch: {ex.Message}", ex);
                _statusManager.Post("Critical AutoRun Error! Check Logs.", MessageType.Error);
                _uiManager.UpdateStatusRight("AutoRun: FAILED (Timer Error)");
                UpdateAutoRunButtonAndStatus(dailyCheckTimer.Enabled, true, "AutoRun: FAILED (Timer Error)");
                autoRunResult = AutoRunActionResult.CriticalError;
            }
            finally
            {
                if (originallyEnabled && autoRunResult != AutoRunActionResult.CriticalError) dailyCheckTimer.Start();
                if (autoRunResult == AutoRunActionResult.ActionAttempted || autoRunResult == AutoRunActionResult.CriticalError)
                {
                    _uiManager.ResetUIOnError(oneClickProcessButton.Visible ? "Generate, Process & Email Report" : "Create Report", _reportPathService.IsEssentialPathConfigurationValid(), File.Exists(_lastGeneratedRawReportPath), File.Exists(_lastGeneratedAnalysisFilePath), IsAnyDailySelected(), dailyCheckTimer.Enabled, autoRunResult == AutoRunActionResult.CriticalError, _uiManager.GetAutoRunStatusLabelText());
                }
            }
        }

        /// <summary>
        /// Handles the Click event for the auto-run toggle button.
        /// </summary>
        private void toggleAutoRunButton_Click(object sender, EventArgs e)
        {
            dailyCheckTimer.Enabled = !dailyCheckTimer.Enabled;
            string statusText = _uiManager.GetAutoRunStatusLabelText();
            bool isAutoRunCompletedForToday = (statusText.Contains("Completed", StringComparison.OrdinalIgnoreCase)) || (statusText.Contains("Done for", StringComparison.OrdinalIgnoreCase)) || (statusText.Contains("FAILED", StringComparison.OrdinalIgnoreCase));
            UpdateAutoRunButtonAndStatus(dailyCheckTimer.Enabled, isAutoRunCompletedForToday, string.Empty);
            Logger.LogInfo($"AutoRun timer {(dailyCheckTimer.Enabled ? "Enabled" : "Disabled")} by user.");
        }
        #endregion

        #region Helper Methods
        /// <summary>
        /// Gathers all parameters for a manual report run from the UI controls.
        /// </summary>
        private ManualReportParameters GatherManualReportParameters() => new() { StartDate = startDatePicker.Value, EndDate = endDatePicker.Value, ReportType = GetSelectedReportType(), FinancialYear = financialYearComboBox.Visible ? financialYearComboBox.SelectedItem?.ToString() : null, IsFemiOnlyChecked = sendToFemiOnlyCheckBox.Checked && sendToFemiOnlyCheckBox.Visible, SkipEmail = skipEmailCheckBox.Checked, IncludeLeadTimeAnalysis = chkIncludeLeadTimeAnalysis.Checked, ReportBaseName = "EstimateSuccessReport", IsDebugBuild = IsDebug };

        /// <summary>
        /// Gets the currently selected ReportType from the ComboBox.
        /// </summary>
        private ReportType GetSelectedReportType() { string? selectedText = null; UIManager.SafeControlUpdate(reportTypeComboBox, () => selectedText = reportTypeComboBox.SelectedItem?.ToString() ?? reportTypeComboBox.Text); return ReportTypeHelper.FromString(selectedText); }

        /// <summary>
        /// Updates the UI to show either the 1-Click button or the two-step buttons.
        /// </summary>
        private void Update1ClickProcessingModeUI() { bool oneClickEnabled = enable1ClickProcessingToolStripMenuItem.Checked; UIManager.SafeControlUpdate(oneClickProcessButton, () => oneClickProcessButton.Visible = oneClickEnabled); UIManager.SafeControlUpdate(createReportButton, () => createReportButton.Visible = !oneClickEnabled); UIManager.SafeControlUpdate(processEmailButton, () => processEmailButton.Visible = !oneClickEnabled); if (oneClickEnabled && oneClickProcessButton != null) UIManager.SafeControlUpdate(oneClickProcessButton, () => oneClickProcessButton.BringToFront()); }

        /// <summary>
        /// Populates the financial year dropdown based on the current date.
        /// </summary>
        private void PopulateFinancialYearDropdown() { UIManager.SafeControlUpdate(financialYearComboBox, () => { string? previouslySelected = financialYearComboBox.SelectedItem?.ToString(); financialYearComboBox.Items.Clear(); string currentFY = _excelProcessor.GetCurrentFinancialYear(true); if (!string.IsNullOrEmpty(currentFY)) { financialYearComboBox.Items.Add(currentFY); string? previousFY = _excelProcessor.GetPreviousFinancialYear(currentFY); if (!string.IsNullOrEmpty(previousFY)) financialYearComboBox.Items.Add(previousFY); } else { financialYearComboBox.Items.Add("FY Unknown"); } if (!string.IsNullOrEmpty(previouslySelected) && financialYearComboBox.Items.Contains(previouslySelected)) { financialYearComboBox.SelectedItem = previouslySelected; } else if (financialYearComboBox.Items.Count > 0) { financialYearComboBox.SelectedIndex = 0; } }); }

        /// <summary>
        /// Validates that the start date is not after the end date.
        /// </summary>
        private bool ValidateInputDates() { if (startDatePicker.Value.Date > endDatePicker.Value.Date) { FlexibleMessageBox.Show(this, "The 'From' date cannot be after the 'To' date.", "Date Range Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return false; } return true; }

        /// <summary>
        /// Validates that the selected date range aligns with the selected financial year.
        /// </summary>
        private bool ValidateFinancialYearSelection() { if (!financialYearComboBox.Visible || financialYearComboBox.SelectedItem == null) return true; string selectedFinYear = financialYearComboBox.SelectedItem.ToString()!; if (!_excelProcessor.IsFinancialYearValid(selectedFinYear, startDatePicker.Value, endDatePicker.Value)) { DialogResult fdr = FlexibleMessageBox.Show(this, $"Date range ({startDatePicker.Value:d} - {endDatePicker.Value:d}) not in Financial Year ({selectedFinYear}).\nContinue?", "FY Mismatch Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning); return fdr == DialogResult.Yes; } return true; }

        /// <summary>
        /// Checks if one of the daily report types is currently selected.
        /// </summary>
        private bool IsAnyDailySelected() { ReportType selectedType = GetSelectedReportType(); return selectedType == ReportType.Daily || selectedType == ReportType.Daily5Day1k; }

        /// <summary>
        /// Re-initializes components and UI text that depend on configuration values.
        /// </summary>
        private void ReinitializeConfigurableComponents() { _appName = _configuration.GetValue<string>(AppConfigKeys.ApplicationInfo.AppName, "QCRA")!; _appVersion = _configuration.GetValue<string>(AppConfigKeys.ApplicationInfo.AppVersion, "1.0.0")!; this.Text = $"{_appName} - {(IsDebug ? "DEBUG" : "RELEASE")} - v{_appVersion}"; _currentAutoRunHour = _configuration.GetValue<int>(AppConfigKeys.AutoRunProcess.CheckHour, 8); _uiManager.SetAutoRunHour(_currentAutoRunHour); UpdateAutoRunButtonAndStatus(dailyCheckTimer.Enabled, false, $"Auto Run: {(dailyCheckTimer.Enabled ? $"Enabled (Next check ~{_currentAutoRunHour}:00)" : "Disabled")}"); bool configIsValid = _reportPathService.IsEssentialPathConfigurationValid(); _uiManager.ResetButtonStatesAfterTypeChange(configIsValid); Update1ClickProcessingModeUI(); }
        #endregion
    }
}