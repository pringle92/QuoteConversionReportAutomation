using OfficeOpenXml;
using QuoteConversionReportAutomation.Interfaces;
using QuoteConversionReportAutomation.Models.Status;
using QuoteConversionReportAutomation.Orchestrators.Interfaces;
using QuoteConversionReportAutomation.Services.Excel;
using QuoteConversionReportAutomation.Services.Interfaces;
using QuoteConversionReportAutomation.Services.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace QuoteConversionReportAutomation.Orchestrators
{
    public class RetrospectiveAnalysisOrchestrator : IRetrospectiveAnalysisOrchestrator
    {
        private readonly IStatusManagerService _statusManager;
        private readonly ExcelCopyData _excelProcessor;

        public RetrospectiveAnalysisOrchestrator(IStatusManagerService statusManager, ExcelCopyData excelProcessor)
        {
            _statusManager = statusManager;
            _excelProcessor = excelProcessor;
        }

        public async Task GenerateAnalysisAsync(string targetFolder, string fileNamePattern, CancellationToken cancellationToken)
        {
            _statusManager.Post("Starting retrospective analysis...", MessageType.InProgress);
            string summaryFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Retrospective_Summary_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            var allLeadTimeData = new List<ExcelCopyData.LeadTimeRecord>();

            try
            {
                var reportFiles = Directory.EnumerateFiles(targetFolder, fileNamePattern, SearchOption.AllDirectories).ToList();
                if (!reportFiles.Any())
                {
                    _statusManager.Post($"No files matching '{fileNamePattern}' found in the selected folder.", MessageType.Warning, TimeSpan.FromSeconds(10));
                    return;
                }

                var sortedFiles = reportFiles.OrderBy(f => Path.GetFileName(f)).ToList();
                _statusManager.Post($"Found {sortedFiles.Count} files. Starting processing...", MessageType.InProgress);

                await Task.Run(() =>
                {
                    for (int i = 0; i < sortedFiles.Count; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        string filePath = sortedFiles[i];
                        _statusManager.Post($"Processing file {i + 1} of {sortedFiles.Count}: {Path.GetFileName(filePath)}", MessageType.InProgress);
                        allLeadTimeData.AddRange(_excelProcessor.ExtractLeadTimeRecords(filePath));
                    }
                }, cancellationToken);

                if (allLeadTimeData.Any())
                {
                    _statusManager.Post("Generating final summary spreadsheet...", MessageType.InProgress);
                    using var summaryPackage = new ExcelPackage();
                    var summaryWorksheet = summaryPackage.Workbook.Worksheets.Add("Lead Time Summary");
                    _excelProcessor.GenerateSummarySheet(summaryWorksheet, allLeadTimeData);

                    if (File.Exists(summaryFilePath)) File.Delete(summaryFilePath);
                    await summaryPackage.SaveAsAsync(new FileInfo(summaryFilePath));
                    _statusManager.Post($"Summary created successfully on your desktop!", MessageType.Success, TimeSpan.FromSeconds(10));
                }
                else
                {
                    _statusManager.Post("No valid lead time records were found in the processed files.", MessageType.Warning, TimeSpan.FromSeconds(10));
                }
            }
            catch (OperationCanceledException) { _statusManager.Post("Analysis cancelled.", MessageType.Warning); }
            catch (Exception ex)
            {
                _statusManager.Post($"An error occurred: {ex.Message}", MessageType.Error);
                Logger.LogError($"Retrospective analysis failed: {ex.Message}", ex);
            }
        }
    }
}