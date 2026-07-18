using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Excel;
using Wengge.OfficeWorker.OpenXml;
using Wengge.OfficeWorker.Office;
using Wengge.OfficeWorker.Presentation;
using Wengge.OfficeWorker.Protocol;
using Wengge.OfficeWorker.Word;
using Wengge.OfficeWorker.Wps;

namespace Wengge.OfficeWorker.Runtime;

public sealed class OfficeWorkerHost : IDisposable
{
    public const int ProtocolVersion = 2;
    private readonly StaDispatcher sta;
    private readonly OwnedProcessJob ownedProcesses;
    private readonly ExcelSessionService excelSessions;
    private readonly ExcelRangeService excelRanges;
    private readonly ExcelWorkbookService excelWorkbooks;
    private readonly ExcelFormulaService excelFormulas;
    private readonly ExcelVbaService excelVba;
    private readonly ExcelUiService excelUi;
    private readonly OpenXmlFileService openXmlFiles;
    private readonly OpenXmlLayoutService openXmlLayout;
    private readonly OpenXmlTableService openXmlTables;
    private readonly OpenXmlDocumentParserService openXmlDocumentParser;
    private readonly WordService word;
    private readonly PresentationService presentation;
    private readonly WpsJsaService wpsJsa;
    private readonly OfficeActionService officeActions;
    private readonly OfficeDocumentService officeDocuments;
    private readonly OfficeObjectService officeObjects;
    private readonly OpenXmlActionService openXmlActions;
    private readonly OfficeSmokeService officeSmoke;

    private OfficeWorkerHost()
    {
        sta = new StaDispatcher();
        ownedProcesses = new OwnedProcessJob();
        var applications = new OfficeApplicationProvider(ownedProcesses);
        excelSessions = new ExcelSessionService(applications);
        excelRanges = new ExcelRangeService(excelSessions);
        excelWorkbooks = new ExcelWorkbookService(excelSessions);
        excelFormulas = new ExcelFormulaService(excelSessions);
        excelVba = new ExcelVbaService(excelSessions);
        excelUi = new ExcelUiService(excelSessions);
        openXmlFiles = new OpenXmlFileService();
        openXmlLayout = new OpenXmlLayoutService();
        openXmlTables = new OpenXmlTableService();
        openXmlDocumentParser = new OpenXmlDocumentParserService();
        officeDocuments = new OfficeDocumentService();
        word = new WordService(applications, officeDocuments);
        presentation = new PresentationService(applications, officeDocuments);
        wpsJsa = new WpsJsaService(applications);
        officeActions = new OfficeActionService(
            new ExcelActionService(
                applications,
                new ExcelQueryActionService(applications),
                new ExcelWorkbookActionService(
                    new ExcelChartActionService(applications),
                    new ExcelObjectActionService(applications),
                    new ExcelTemplatePrintActionService(applications)),
                new ExcelFormulaActionService(applications),
                new ExcelCrossOfficeActionService(applications)),
            new WordActionService(
                new WordFormattingActionService(applications),
                new WordReferenceRevisionActionService(applications, officeDocuments),
                new WordMailMergeContentActionService(applications),
                new WordLinkedContentActionService(applications)),
            new PresentationActionService(
                new PresentationEditActionService(applications),
                new PresentationInspectionActionService(applications),
                new PresentationBrandingActionService(applications),
                new PresentationLinkedContentActionService(applications),
                new PresentationPlaybackActionService(applications)));
        officeObjects = new OfficeObjectService(officeDocuments);
        officeSmoke = new OfficeSmokeService(officeDocuments, excelSessions);
        openXmlActions = new OpenXmlActionService(
            new OpenXmlExcelActionService(openXmlTables),
            new OpenXmlWordActionService(openXmlTables),
            new OpenXmlPresentationActionService());
    }

    public static OfficeWorkerHost Create() => new();

    public Task<object?> DispatchAsync(RpcRequest request, CancellationToken cancellationToken)
    {
        var parameters = request.Params;
        return request.Method switch
        {
            "worker.health" => Task.FromResult<object?>(new
            {
                ready = true,
                protocolVersion = ProtocolVersion,
                workerVersion = typeof(OfficeWorkerHost).Assembly.GetName().Version?.ToString(),
                runtime = Environment.Version.ToString(),
                architecture = System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture.ToString().ToLowerInvariant(),
            }),
            "office.action.execute" => OnSta(() => officeActions.Execute(OfficeActionRequest.Parse(parameters)), cancellationToken),
            "openxml.action.execute" => Task.Run<object?>(() => openXmlActions.Execute(OfficeActionRequest.Parse(parameters)), cancellationToken),
            "openxml.parseDocument" => Task.Run<object?>(() => openXmlDocumentParser.Parse(parameters.RequiredString("filePath")), cancellationToken),
            "office.documents.list" => OnSta(() => officeDocuments.ListDocuments(parameters.OptionalString("app")), cancellationToken),
            "office.documents.activate" => OnSta(() => officeDocuments.ActivateDocument(parameters), cancellationToken),
            "office.objects.list" => OnSta(() => officeObjects.ListObjects(parameters), cancellationToken),
            "office.objects.activate" => OnSta(() => officeObjects.ActivateObject(parameters), cancellationToken),
            "office.transaction.prepare" => OnSta(() => officeDocuments.PrepareTransaction(parameters.PropertyOrEmpty("filePaths")), cancellationToken),
            "office.transaction.restoreFiles" => OnSta(() => officeDocuments.RestoreTransactionFiles(parameters.PropertyOrEmpty("files")), cancellationToken),
            "office.smoke.markWordBookmarkDirty" => OnSta(() => officeSmoke.MarkWordBookmarkDirty(parameters), cancellationToken),
            "office.smoke.openFixtures" => OnSta(() => officeSmoke.OpenFixtures(parameters), cancellationToken),
            "office.smoke.closeFixtures" => OnSta(() => officeSmoke.CloseFixtures(parameters), cancellationToken),
            "office.smoke.listProcesses" => Task.Run<object?>(officeSmoke.ListProcesses, cancellationToken),
            "office.smoke.runningProcesses" => Task.Run<object?>(() => officeSmoke.RunningProcesses(parameters.PropertyOrEmpty("ids")), cancellationToken),
            "office.smoke.excel.getDisplayAlerts" => OnSta(officeSmoke.GetExcelDisplayAlerts, cancellationToken),
            "office.smoke.excel.setDisplayAlerts" => OnSta(() => officeSmoke.SetExcelDisplayAlerts(parameters), cancellationToken),
            "office.smoke.excel.setStructureProtected" => OnSta(() => officeSmoke.SetExcelStructureProtected(parameters), cancellationToken),
            "excel.detectStatus" => OnSta(excelSessions.DetectStatus, cancellationToken),
            "excel.connect" => OnSta(excelSessions.Connect, cancellationToken),
            "excel.selectHost" => OnSta(() => excelSessions.SelectHost(parameters.RequiredString("host")), cancellationToken),
            "excel.workbook.inspect" => OnSta(excelWorkbooks.Inspect, cancellationToken),
            "excel.workbook.open" => OnSta(() => excelWorkbooks.Open(parameters.RequiredString("filePath")), cancellationToken),
            "excel.workbook.create" => OnSta(() => excelWorkbooks.Create(parameters.RequiredString("filePath"), parameters.PropertyOrEmpty("sheetNames")), cancellationToken),
            "excel.workbook.save" => OnSta(() => excelWorkbooks.Save(parameters.OptionalString("saveAsPath")), cancellationToken),
            "excel.workbook.switch" => OnSta(() => excelWorkbooks.Switch(parameters.RequiredString("workbookName")), cancellationToken),
            "excel.range.read" => OnSta(() => excelRanges.Read(parameters.RequiredString("sheetName"), parameters.RequiredString("range"), parameters.OptionalString("expand") ?? "none"), cancellationToken),
            "excel.range.write" => OnSta(() => excelRanges.Write(
                parameters.RequiredString("sheetName"),
                parameters.RequiredString("range"),
                parameters.PropertyOrEmpty("values"),
                parameters.OptionalBoolean("legacyCse")), cancellationToken),
            "excel.range.clear" => OnSta(() => excelRanges.Clear(parameters.RequiredString("sheetName"), parameters.RequiredString("range")), cancellationToken),
            "excel.selection.read" => OnSta(() => excelRanges.GetSelection(includeValues: true), cancellationToken),
            "excel.selection.address" => OnSta(() => excelRanges.GetSelection(includeValues: false), cancellationToken),
            "excel.formula.context" => OnSta(() => excelFormulas.GetContext(parameters.RequiredString("sheetName"), parameters.OptionalString("range")), cancellationToken),
            "excel.vba.detect" => OnSta(excelVba.DetectCapabilities, cancellationToken),
            "excel.vba.run" => OnSta(() => excelVba.RunMacro(parameters.RequiredString("macroName"), parameters.PropertyOrEmpty("args")), cancellationToken),
            "excel.vba.writeModule" => OnSta(() => excelVba.WriteModule(
                parameters.RequiredString("moduleName"),
                parameters.RequiredString("code"),
                parameters.OptionalString("entryPoint"),
                parameters.OptionalBoolean("save"),
                parameters.OptionalString("saveAsPath")), cancellationToken),
            "excel.ui.addControl" => OnSta(() => excelUi.AddControl(parameters), cancellationToken),
            "excel.ui.removeControl" => OnSta(() => excelUi.RemoveControl(parameters.RequiredString("sheetName"), parameters.RequiredString("name")), cancellationToken),
            "excel.ui.listControls" => OnSta(() => excelUi.ListControls(parameters.RequiredString("sheetName")), cancellationToken),
            "excel.ui.createForm" => OnSta(() => excelUi.CreateForm(parameters), cancellationToken),
            "excel.ui.addMenu" => OnSta(() => excelUi.AddMenu(parameters), cancellationToken),
            "excel.sheet.operation" => OnSta(() => excelWorkbooks.SheetOperation(parameters.RequiredString("operation"), parameters.RequiredString("sheetName"), parameters.PropertyOrEmpty("options")), cancellationToken),
            "word.open" => OnSta(() => word.Open(parameters.RequiredString("filePath")), cancellationToken),
            "word.detectStatus" => OnSta(word.DetectStatus, cancellationToken),
            "word.inspect" => OnSta(word.Inspect, cancellationToken),
            "word.readText" => OnSta(() => word.ReadText(parameters.OptionalInt32("maxChars", 12_000)), cancellationToken),
            "word.insertText" => OnSta(() => word.InsertText(parameters.RequiredString("text"), parameters.OptionalString("position") ?? "end"), cancellationToken),
            "word.insertHeading" => OnSta(() => word.InsertHeading(parameters.RequiredString("text"), parameters.OptionalInt32("level", 1), parameters.OptionalString("position") ?? "end"), cancellationToken),
            "word.replaceText" => OnSta(() => word.ReplaceText(parameters.RequiredString("findText"), parameters.OptionalString("replaceText") ?? string.Empty, parameters.OptionalBoolean("matchCase")), cancellationToken),
            "word.save" => OnSta(() => word.Save(parameters.OptionalString("saveAsPath")), cancellationToken),
            "presentation.open" => OnSta(() => presentation.Open(parameters.RequiredString("filePath")), cancellationToken),
            "presentation.detectStatus" => OnSta(presentation.DetectStatus, cancellationToken),
            "presentation.inspect" => OnSta(presentation.Inspect, cancellationToken),
            "presentation.readSlide" => OnSta(() => presentation.ReadSlide(parameters.OptionalInt32("slideIndex", 1)), cancellationToken),
            "presentation.addSlide" => OnSta(() => presentation.AddSlide(parameters.OptionalString("title"), parameters.OptionalString("body"), parameters.OptionalString("layout") ?? "titleAndContent"), cancellationToken),
            "presentation.setShapeText" => OnSta(() => presentation.SetShapeText(parameters.OptionalInt32("slideIndex", 1), parameters.RequiredString("text"), parameters.OptionalString("shapeName"), parameters.OptionalInt32("shapeIndex", 1)), cancellationToken),
            "presentation.replaceText" => OnSta(() => presentation.ReplaceText(parameters.RequiredString("findText"), parameters.OptionalString("replaceText") ?? string.Empty, parameters.OptionalBoolean("matchCase")), cancellationToken),
            "presentation.save" => OnSta(() => presentation.Save(parameters.OptionalString("saveAsPath")), cancellationToken),
            "wps.jsa.detect" => wpsJsa.DetectAsync(cancellationToken),
            "wps.jsa.write" => wpsJsa.WriteAsync(parameters, cancellationToken),
            "openxml.inspect" => Task.Run<object?>(() => openXmlFiles.Inspect(parameters.RequiredString("filePath")), cancellationToken),
            "openxml.replaceText" => Task.Run<object?>(() => openXmlFiles.ReplaceText(
                parameters.RequiredString("filePath"),
                parameters.RequiredString("findText"),
                parameters.OptionalString("replaceText") ?? string.Empty,
                parameters.OptionalString("outputPath"),
                parameters.OptionalBoolean("matchCase")), cancellationToken),
            "openxml.inspectLayout" => Task.Run<object?>(() => openXmlLayout.Inspect(parameters.RequiredString("filePath"), parameters.OptionalString("target")), cancellationToken),
            "openxml.inspectTable" => Task.Run<object?>(() => openXmlTables.Inspect(parameters.RequiredString("filePath"), parameters.OptionalString("target")), cancellationToken),
            "openxml.applyTableStyle" => Task.Run<object?>(() => openXmlTables.ApplyStyle(
                parameters.RequiredString("filePath"),
                parameters.OptionalString("style") ?? "professional",
                parameters.OptionalString("outputPath"),
                parameters.OptionalString("target")), cancellationToken),
            "openxml.snapshot" => Task.FromResult<object?>(new
            {
                engine = "openxml",
                operation = "snapshot",
                documentType = Path.GetExtension(parameters.RequiredString("filePath")).TrimStart('.'),
                filePath = parameters.RequiredString("filePath"),
                outputPath = parameters.OptionalString("outputPath"),
                target = parameters.OptionalString("target"),
                supported = false,
                error = "Open XML 不负责像素渲染，请使用 COM 快照能力",
            }),
            _ => throw new OfficeWorkerException("method_not_found", $"未知 Worker 方法: {request.Method}"),
        };
    }

    public void Dispose()
    {
        officeSmoke.Dispose();
        wpsJsa.Dispose();
        sta.Dispose();
        ownedProcesses.Dispose();
    }

    private async Task<object?> OnSta(Func<object> operation, CancellationToken cancellationToken) =>
        await sta.InvokeAsync(operation, cancellationToken);
}
