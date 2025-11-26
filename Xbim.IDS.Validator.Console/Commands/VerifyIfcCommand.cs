using ARXCommand;
using ClosedXML.Excel;
using ARXCommand;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System.CommandLine.Invocation;
using System.Diagnostics;
using Xbim.Common;
using Xbim.IDS.Validator.Common;
using Xbim.IDS.Validator.Console.Internal;
using Xbim.IDS.Validator.Core;
using Xbim.IDS.Validator.Core.Interfaces;
using Xbim.Ifc4.Interfaces;
using Xbim.InformationSpecifications;
using Xbim.IO.CobieExpress;
using Xbim.IO.Memory;
using static System.ConsoleColor;
using static Xbim.IDS.Validator.Console.CliOptions;

namespace Xbim.IDS.Validator.Console.Commands
{
    /// <summary>
    /// Verifies IFC and COBie models against a set of IDS files.
    /// </summary>
    internal class VerifyIfcCommand : ICommand
    {
        private readonly IIdsModelValidator idsValidator;
        private readonly ILogger<VerifyIfcCommand> logger;
        private readonly IdsConfig config;
        private ConsoleLogger console = new ConsoleLogger(Verbosity.Normal);

        /// <summary>
        /// Constructs a new <see cref="VerifyIfcCommand"/>
        /// </summary>
        /// <param name="validator"></param>
        /// <param name="logger"></param>
        /// <param name="config"></param>
        public VerifyIfcCommand(IIdsModelValidator validator, ILogger<VerifyIfcCommand> logger, IOptions<IdsConfig> config)
        {
            idsValidator = validator;
            this.logger = logger;
            this.config = config.Value;
        }

        public async Task<int> ExecuteAsync(InvocationContext ctx)
        {
            var idsFiles = ctx.ParseResult.GetValueForOption(VerifyCommand.IdsFilesOption);
            var modelFiles = ctx.ParseResult.GetValueForArgument(VerifyCommand.ModelFilesArgument);
            var verbosity = ctx.ParseResult.GetValueForOption(VerbosityOption);
            var specNameFilter = ctx.ParseResult.GetValueForOption(VerifyCommand.IdsFilterOption);

            return await Execute(idsFiles, modelFiles, verbosity, specNameFilter);
        }


        private async Task<int> Execute(string[] idsFiles, string[] modelFiles, Verbosity verbosity, string specNameFilter)
        {
            var failedSpecs = 0;
            console = new ConsoleLogger(verbosity);

            foreach (var modelFile in modelFiles)
            {
                IModel? model = default;
                try
                {
                    if (File.Exists(modelFile) != true)
                    {
                        logger.LogWarning("Model {file} not found", modelFile);
                        return -1;
                    }

                    console.WriteImportantLine("IFC File: {0}", modelFile);
                    console.WriteInfoLine("Loading Model...");
                    var sw = new Stopwatch();
                    sw.Start();
#if SQLite

                IModel model = BuildModelSqlLite(modelFile);
                var ifcFlexDb = model as IfcFlexDb;
                using var tc = ifcFlexDb!.BeginTypeCaching();
                using var inverseCache = ifcFlexDb.BeginInverseCaching();
                using var entityeCache = ifcFlexDb.BeginEntityCaching();
                OptimiseActivationStrategy(ifcFlexDb);
#else
                    model = BuildModel(modelFile);
                    using var ic = model.BeginInverseCaching();
                    using var ec = model.BeginEntityCaching();
#endif
                    console.WriteInfoLine("Model loaded in {0}s", sw.Elapsed.TotalSeconds);
                    // Normally we'd inject rather than service discovery

                    string schemaVersion = model.SchemaVersion.ToString().ToUpper();


                    foreach (var ids in idsFiles)
                    {
                        if (File.Exists(ids) != true)
                        {
                            logger.LogWarning("IDS {file} not found", ids);
                            return -2;
                        }
                        console.WriteImportantLine("IDS File: {0}", ids);
                        console.WriteInfoLine("Validating...");
                       
                        var options = new VerificationOptions
                        {
                            IncludeSubtypes = false,
                            OutputFullEntity = true,
                            AllowDerivedAttributes = false,
                            PerformInPlaceSchemaUpgrade = true,
                            PermittedIdsAuditStatuses = VerificationOptions.AnyState,
                            SkipIncompatibleSpecification = true,
                            SpecExecutionFilter = s => 
                                string.IsNullOrEmpty(specNameFilter) ||
                                (!specNameFilter.StartsWith("*") && (s.Guid.StartsWith(specNameFilter) || s.Name?.StartsWith(specNameFilter) == true)) || // Spec ID starts with pattern
                                (specNameFilter.StartsWith("*") && s?.Name?.Contains(specNameFilter.Replace("*", ""), StringComparison.InvariantCultureIgnoreCase) != false), // Spec Name contains pattern
                            RuntimeTokens = config.Detokenise ? config.Tokens : new(),
                        };

                        if(model is CobieModel)
                        {
                            options.IncludeSubtypes = true;
                            options.AllowDerivedAttributes = true;
                        }

                        var results = await idsValidator.ValidateAgainstIdsAsync(model, ids, logger, OutputRequirement, options);



                        sw.Stop();
                        failedSpecs += results.ExecutedRequirements.Count(r => r.Status == ValidationStatus.Fail);

                        WriteSummary(results, sw, Path.GetFileName(modelFile), Path.GetFileName(ids));

                        if (results.Status == ValidationStatus.Error)
                        {
                            console.WriteImportantLine($"Validation failed to run: {results.Message}");
                            return -1;
                        }

                        if (results.Status == ValidationStatus.Inconclusive)
                        {
                            console.WriteImportantLine($"Validation failed to run: {results.Message}");
                            return -1;
                        }


                        //string curFolder = Path.GetDirectoryName(modelFile) + "\\"+ Path.GetFileNameWithoutExtension(modelFile)+ "\\";

                        string curFolder = Path.GetDirectoryName(modelFile) + "\\";




                        string reportFolder = curFolder+Path.GetFileNameWithoutExtension(modelFile)+"\\";


                        if (!Directory.Exists(reportFolder))
                        {
                            // Try to create the directory.
                            DirectoryInfo di = Directory.CreateDirectory(curFolder);
                        }


                        //*****stefano


                        string txtReportPath = reportFolder + Path.ChangeExtension(Path.GetFileNameWithoutExtension(modelFile)+"_Report-IDS", ".txt");
                        string htmlReportPath = reportFolder+Path.ChangeExtension(Path.GetFileNameWithoutExtension(modelFile)+"_Report-IDS", ".html");
                        string PDFReportPath = reportFolder+Path.ChangeExtension(Path.GetFileNameWithoutExtension(modelFile)+"_Report-IDS", ".pdf");
                        string nomeFileExcel = reportFolder+Path.ChangeExtension(Path.GetFileNameWithoutExtension(modelFile)+"_Report-IDS", ".xlsx");
                        string rtfReportPath = reportFolder + Path.ChangeExtension(Path.GetFileNameWithoutExtension(modelFile)+"_Report-IDS", ".rtf");
                        string wordReportPath = reportFolder + Path.ChangeExtension(Path.GetFileNameWithoutExtension(modelFile)+"_Report-IDS", ".docx");
                        string bcfreportPath = reportFolder + Path.ChangeExtension(Path.GetFileNameWithoutExtension(modelFile) + "_Report-IDS", ".bcfzip");


                        var ARXNewCommand = new ARXNewCommand();
                        //ARXNewCommand.BCFExport(results, modelFile, bcfreportPath);
                        //await SaveResultsToFileAsync(results, txtReportPath);

                        //SaveResultsToRTF(results, rtfReportPath);

                        ARXNewCommand.SaveResultsToWORD(results, wordReportPath, Path.GetFileName(modelFile), Path.GetFileName(ids), model.SchemaVersion.ToString());

                        {
                            if (results.Status is ValidationStatus.Fail)

                                ARXNewCommand.SaveExcelFile(results, nomeFileExcel);

                        }



                        var reportData = new IdsReportMapper().Map(results, Path.GetFileName(modelFile), model.SchemaVersion.ToString());

                        reportData.SpecificationResults.OrderByDescending(r => r.Status);



                        //// 4. Genera il file HTML.
                        var generator = new IdsReportGenerator.HtmlReportGenerator();

                        generator.Generate(reportData, htmlReportPath);

                       

                        
                    }

                }
                finally
                {
                    model?.Dispose();
                }
            }
            return failedSpecs;
        }




        private void WriteSummary(ValidationOutcome results, Stopwatch sw, string ifcFile, string idsFile)
        {
            console?
                .WriteInfo(Blue, "IDS Validation summary\n")
                .WriteInfo(Gray, $"IDS:   ")
                .WriteInfoLine(White, idsFile)
                .WriteInfo(Gray, "Model: ")
                .WriteInfoLine(White, ifcFile)
                .WriteInfoLine(Gray, "---------------------------------------------------------------------------")
                ;
            var totalRun = results.ExecutedRequirements.Count;
            var totalPass = results.ExecutedRequirements.Count(r => r.Status == ValidationStatus.Pass);
            var totalInconclusive = results.ExecutedRequirements.Count(r => r.Status == ValidationStatus.Inconclusive);
            var totalSkipped = results.ExecutedRequirements.Count(r => r.Status == ValidationStatus.Skipped);
            var totalFail = results.ExecutedRequirements.Count(r => r.Status == ValidationStatus.Fail);
            var totalError = results.ExecutedRequirements.Count(r => r.Status == ValidationStatus.Error);

            var totalElementsTested = results.ExecutedRequirements.Sum(r => r.ApplicableResults.Count(r => r.ValidationStatus != ValidationStatus.Error));
            var totalPassedResults = results.ExecutedRequirements.Sum(r => r.PassedResults.Count());
            var totalFailedResults = results.ExecutedRequirements.Sum(r => r.FailedResults.Count());
            var totalPercent = totalElementsTested > 0 ? (float)totalPassedResults / totalElementsTested * 100 : 0;
            
            
            
            console?.WriteInfoLine(White, $"Detailed Results:");
            console?.WriteImportantLine(Gray, " no      Pass /Total  %age   #Fail  Specification");
            console?.WriteImportantLine(Gray, "------ -------------- ------ -----  ---------------------------------------------");
            int i = 1;

            foreach (var req in results.ExecutedRequirements)
            {
                int? passed = req.PassedResults.Count();
                passed = passed == 0 ? null : passed;
                int? failed = req.FailedResults.Count();
                failed = failed == 0 ? null : failed;
                var total = req.ApplicableResults.Count(r => r.ValidationStatus != ValidationStatus.Error);
                float? percent = req.ApplicableResults.Count() > 0 ? (float)(passed ?? 0) / req.ApplicableResults.Count() * 100 : 0;
                if (req.Status == ValidationStatus.Pass && req.ApplicableResults.Count() == 0 || req.Status == ValidationStatus.Skipped)
                {
                    percent = null;
                }
                if(console?.MinVerbosity <= Verbosity.Minimal && req.ApplicableResults.Any() == false && req.Status == ValidationStatus.Pass && req.Specification.Cardinality.NoMatchingEntities == false )
                {
                    continue;   // Skip 'no applicable rows' for quieter log levels for more consise logs 
                }
                console?.WriteImportant(White, $"{i++,4}");
                console?.WriteImportant(console.GetColorForStatus(req.Status), $"{StatusIcon(req.Status)}");
                console?.WriteImportant(White, $" [{passed,5}")
                    .WriteImportant(Gray, $" /")
                    .WriteImportant(White, $"{total,5}]")
                //.WriteInfo(Gray, $" passed ")
                //   .WriteInfo(Gray, $" failed ")
                    .WriteImportant(DarkYellow, $" {percent,5:0.0}% ")
                    .WriteImportant(Red, $"{failed,5}")
                    .WriteImportant(Gray, $"  {req.Specification.Name}\n", Gray);
                // [{passed} passed from {req.ApplicableResults.Count}]", 
            }
            console?.WriteImportantLine(Gray, "------ -------------- ------ -----  ---------------------------------------------");
            console?
                .WriteImportant(console.GetColorForStatus(results.Status), "    " + StatusIcon(results.Status))
                .WriteImportant(White, $" [{totalPassedResults,5}")
                .WriteImportant(Gray, $" /")
                .WriteImportant(White, $"{totalElementsTested,5}]")
                .WriteImportant(DarkYellow, $" {totalPercent,5:0.0}% ")
                //.WriteInfo(Gray, $" tested ")
                .WriteImportant(Red, $"{totalFailedResults,5}  ")
                //.WriteInfo(Gray, $" failed ")
                ;

            console?
            .WriteImportant(White, $"Specifications Run: {totalRun} ")
            .WriteImportant(Green, $"Pass: {totalPass} ")
            .WriteImportant(Red, $"Fail: {totalFail} ")
            .WriteImportant(Cyan, $"Not Run: {totalSkipped} ")
            .WriteImportant(Yellow, $"Incomplete: {totalInconclusive} ")
            .WriteImportant(DarkRed, $"Error: {totalError}")
            .WriteImportant(DarkGreen, $" in {sw.Elapsed.TotalSeconds} secs\n");
            console?.WriteImportantLine(White, "===========================================================================\n");
            console?.WriteImportantLine(White, $"{ifcFile} against {idsFile}");
        }

        private Task OutputRequirement(ValidationRequirement req,SpecificationsGroup specificationsGroup)
        {
            var passed = req.PassedResults.Count();
            console.WriteColored(req.Status, req.Status.ToString());

            if (req.Status == ValidationStatus.Fail || req.Status == ValidationStatus.Error)
            {
                console
                    .WriteWarning(Gray, $" : {req.Specification.Name} [{passed} passed from {req.ApplicableResults.Count}]")
                    .WriteWarning(Cyan, $" {req.Specification.Cardinality.Description} Requirement\n");
            }
            else
            {
                console
                    .WriteInfo(Gray, $" : {req.Specification.Name} [{passed} passed from {req.ApplicableResults.Count}]")
                    .WriteInfo(Cyan, $" {req.Specification.Cardinality.Description} Requirement\n");

            }
            console?.WriteDetailLine(Blue, $"  🔎  For {req.Specification.Applicability.GetApplicabilityDescription().SplitClauses()}\n");
            if (req.Specification.Cardinality.AllowsRequirements)
                console?.WriteDetailLine(DarkGreen, $"  📏  It is {req.Specification.Cardinality.Description} that elements {req.Specification.Requirement?.GetRequirementDescription().SplitClauses()}\n");


            foreach (var itm in req.ApplicableResults)
            {
                if (req.Status == ValidationStatus.Error)
                {
                    console?.WriteImportantLine(console.GetColorForStatus(itm.ValidationStatus), "  " + StatusIcon(itm.ValidationStatus));
                    foreach (var msg in itm.Messages.Where(m => m.Status != ValidationStatus.Pass))
                    {
                        console?.WriteImportant(DarkRed, $": {msg?.Reason}\n")
                        .WriteImportant(DarkGray, $"     {msg}\n");
                    }

                }
                else if (req.IsFailure(itm))
                {
                    console?.WriteImportant(console.GetColorForStatus(itm.ValidationStatus), "  " + StatusIcon(itm.ValidationStatus))
                        .WriteImportant(Red, $"{itm.Requirement?.Name} {itm.Requirement?.Description}")
                        .WriteImportant(Gray, $"{itm.FullEntity}\n");
                    foreach (var msg in itm.Messages.Where(m => m.Status != ValidationStatus.Pass))
                    {
                        var msgtxt = msg.ToString()
                        //.Replace("{", "{{")
                        //    .Replace("}", "}}")
                            ;
                        console?.WriteImportant(DarkRed, $"     {msgtxt}\n");
                    }
                }
                else
                {
                    console?.WriteDetail(console.GetColorForStatus(itm.ValidationStatus), "  " + StatusIcon(itm.ValidationStatus))
                        .WriteDetail(DarkGray, $"{itm.Requirement?.Name} {itm.Requirement?.Description}")
                        .WriteDetail(Gray, $"{itm.FullEntity}\n");
                    foreach (var msg in itm.Messages.Where(m => m.Status == ValidationStatus.Pass))
                    {
                        console?.WriteTrace(DarkGray, $"     {msg}\n");
                    }
                }
                //Console.Write(".");
            }
            console?.WriteInfoLine("");
            return Task.CompletedTask;
        }

        private static string StatusIcon(ValidationStatus status)
        {
            return status switch
            {
                ValidationStatus.Pass => "✔️",
                ValidationStatus.Inconclusive => "❓",
                ValidationStatus.Fail => "❌",
                ValidationStatus.Error => "⚠️",
                ValidationStatus.Skipped => "💤",
                _ => throw new NotImplementedException(),
            };
        }


        private static IModel BuildModel(string modelFile)
        {
            if (modelFile.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return OpenCOBie(modelFile);
            }
#if SQLite

        if (modelFile.EndsWith(".db", StringComparison.OrdinalIgnoreCase))
        {
            // return Flex DB
            var flexdb = new IfcFlexDb();
            flexdb.Open(modelFile);
            flexdb.BeginTypeCaching();

            flexdb.AddActivationDepth<IIfcRelDefinesByProperties>(2);
            flexdb.AddActivationDepth<IIfcRelDefinesByType>(2);
            flexdb.AddActivationDepth<IIfcRelAssociatesClassification>(2);
            flexdb.AddActivationDepth<IIfcRelAggregates>(2);

            return flexdb;
        }
#endif

            return MemoryModel.OpenRead(modelFile);
        }

        private static IModel OpenCOBie(string file)
        {
            var mapping = IO.CobieExpress.CobieModel.GetMapping();
            mapping.ClassMappings.RemoveAll(m => m.Class == "System");
            mapping.ClassMappings.RemoveAll(m => m.Class.StartsWith("Attribute"));
            mapping.ClassMappings.RemoveAll(m => m.Class.StartsWith("Zone"));
            //foreach(var map in mapping.ClassMappings)
            //{
            //    Console.WriteLine(map.Class);
            //}
            var model = IO.CobieExpress.CobieModel.ImportFromTable(file, out string report, mapping);

            return model;
        }

        //static string SanitizeSheetName(string name)
        //{
        //    if (string.IsNullOrWhiteSpace(name)) return "Foglio";
        //    var invalidChars = new[] { '\\', '/', '?', '*', '[', ']' };
        //    foreach (char c in invalidChars) { name = name.Replace(c, '_'); }
        //    if (name.Length > 31) { name = name.Substring(0, 31); }
        //    return name;
        //}



        //public void SaveExcelFile(ValidationOutcome outcome, string fileExcel)
        //{

        //    try
        //    {
        //        using var workbook = new XLWorkbook();
        //        string strNameEntity = "";
        //        int Riga = 2;
        //        int Colonna = 1;
        //        var tmpRiga = new strRigaExcel();
        //        var tmpRigaLista = new List<strRigaExcel>();

        //        foreach (var requirement in outcome.ExecutedRequirements)
        //        {
        //            if (requirement.Status is ValidationStatus.Fail)
        //            {
        //                var sheetName = SanitizeSheetName(requirement.Specification.Name);
        //                var worksheet = workbook.Worksheets.Add(sheetName);
        //                worksheet.Cell(1, 1).Value = "ERRORE";
        //                worksheet.Cell(1, 1).Style.Fill.SetBackgroundColor(XLColor.Gray);
        //                worksheet.Cell(1, 4).Value = "NOME OGGETTO";
        //                worksheet.Cell(1, 5).Value = "GUID";
        //                worksheet.Cell(1, 6).Value = "RevitID";
        //                worksheet.Cell(1, 7).Value = "Elemento";


        //                // Dettagli errori
        //                var failedEntities = requirement.ApplicableResults.Where(e => e.ValidationStatus != ValidationStatus.Pass);
        //                if (failedEntities.Any())
        //                {

        //                    foreach (var entity in failedEntities)
        //                    {

        //                        var ifcEntity = entity.FullEntity;
        //                        var ifcRoot = ifcEntity as IIfcRoot;



        //                        tmpRiga.Elemento = entity;
        //                        if (entity is IIfcObject failingObject1)
        //                        {
        //                            tmpRiga.GlobalID = failingObject1.GlobalId;

        //                        }

        //                        foreach (var message in entity.Messages)
        //                        {
        //                            if (message.Status != ValidationStatus.Pass)
        //                            {
        //                                string[] words = ifcRoot?.Name.ToString().Split(":");

        //                                var numWord = words.Count();


        //                                worksheet.Cell(Riga, 1).Value = message.ToString();
        //                                worksheet.Cell(Riga, 1).Style.Alignment.WrapText=true;

        //                                worksheet.Cell(Riga, 4).Value = ifcRoot?.Name.ToString();
        //                                worksheet.Cell(Riga, 5).Value = ifcRoot?.GlobalId.ToString();
        //                                worksheet.Cell(Riga, 6).Value = words[numWord-1];
        //                                worksheet.Cell(Riga, 7).Value = entity.FullEntity.ToString();
        //                                worksheet.Row(Riga).AdjustToContents(1);
        //                                //worksheet.Row(Riga).ClearHeight();


        //                                Riga += 1;

        //                            }


        //                        }


        //                    }

        //                }
        //                Riga = 2;
        //                worksheet.Columns().AdjustToContents();
        //                worksheet.Rows().AdjustToContents(1);
        //            }
        //        }


        //        workbook.SaveAs(fileExcel);
        //        workbook.Dispose();
        //    }
        //    catch (Exception ex)
        //    {
        //        // Cattura altre possibili eccezioni (es. percorso non valido).
        //        //Console.WriteLine($"Si è verificato un errore imprevisto: {ex.Message}");
        //        //Console.WriteLine($"ERRORE: {ex.Message}");
        //        if (ex.InnerException != null)
        //        {
        //            //Console.WriteLine($"Dettaglio: {ex.InnerException.Message}");
        //        }
        //        //Console.WriteLine($"StackTrace: {ex.StackTrace}");
        //    }
        //}
        //public static void SaveResultsToWORD(ValidationOutcome outcome, string outputPath, string modello, string idsFile, string IFCVersion)
        //{
        //    // Crea un nuovo documento Word.
        //    using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
        //    {
        //        // Aggiungi una parte principale al documento.
        //        MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
        //        mainPart.Document = new Document();
        //        Body body = mainPart.Document.AppendChild(new Body());


        //        SectionProperties sectionProps = new SectionProperties();


        //        PageSize pageSize = new PageSize()
        //        {
        //            Width = (UInt32Value)23810U, // Larghezza A3 in Twips (altezza se verticale)
        //            Height = (UInt32Value)16836U  // Altezza A3 in Twips (larghezza se verticale)
        //        };

        //        pageSize.Orient = PageOrientationValues.Landscape;

        //        sectionProps.Append(pageSize);


        //        // Aggiungi la sezione al corpo del documento
        //        body.Append(sectionProps);





        //        RunProperties runProperties1 = new RunProperties(); // Definisce le proprietà del testo (es. colore).
        //        runProperties1.AppendChild(new Color() { Val = "#3381d3" });
        //        FontSize fontSize = new FontSize() { Val = "40" };
        //        runProperties1.Append(fontSize);
        //        Bold bold = new Bold(); // Equivale a <w:b/>
        //        runProperties1.Append(bold);
        //        Paragraph para1 = body.AppendChild(new Paragraph());
        //        Run run1 = para1.AppendChild(new Run());
        //        run1.AppendChild(runProperties1);

        //        run1.AppendChild(new Text($"Nome Modello:   {modello}"));

        //        RunProperties runProperties2 = new RunProperties(); // Definisce le proprietà del testo (es. colore).
        //        runProperties2.AppendChild(new Color() { Val = "#3381d3" });
        //        fontSize = new FontSize() { Val = "40" };
        //        runProperties2.Append(fontSize);
        //        bold = new Bold(); // Equivale a <w:b/>
        //        runProperties2.Append(bold);
        //        Paragraph para2 = body.AppendChild(new Paragraph());
        //        Run run2 = para2.AppendChild(new Run());
        //        run2.AppendChild(runProperties2);

        //        run2.AppendChild(new Text($"Nome File IDS:   {idsFile}"));

        //        body.AppendChild(new Paragraph(new Run(new Text(""))));


        //        RunProperties runProperties3 = new RunProperties(); // Definisce le proprietà del testo (es. colore).
        //        runProperties3.AppendChild(new Color() { Val = "#3381d3" });
        //        fontSize = new FontSize() { Val = "40" };
        //        runProperties3.Append(fontSize);
        //        bold = new Bold(); // Equivale a <w:b/>
        //        runProperties3.Append(bold);
        //        Paragraph para3 = body.AppendChild(new Paragraph());
        //        Run run3 = para3.AppendChild(new Run());
        //        run3.AppendChild(runProperties3);

        //        run3.AppendChild(new Text($"Versione IFC: {IFCVersion} - Data Elaborazione: {DateTime.Now:dd-MM-yyyy HH:mm:ss}"));

        //        body.AppendChild(new Paragraph(new Run(new Text(""))));





        //        foreach (ValidationRequirement req in outcome.ExecutedRequirements)
        //        {
        //            var passed = req.PassedResults.Count();

        //            //run2.AppendChild(new Text(req.Status.ToString()));

        //            if (req.Status == ValidationStatus.Fail || req.Status == ValidationStatus.Error)
        //            {

        //                Paragraph paraN = body.AppendChild(new Paragraph());
        //                Run runN = paraN.AppendChild(new Run());
        //                RunProperties runPropsN = new RunProperties();
        //                // Imposta il colore a nero. Anche se il nero è il colore predefinito, lo si specifica per chiarezza.
        //                fontSize = new FontSize() { Val = "30" };
        //                runPropsN.Append(fontSize);
        //                bold = new Bold(); // Equivale a <w:b/>
        //                runPropsN.Append(bold);
        //                runPropsN.Append(new Color() { Val = "#FF0000" });
        //                runN.Append(runPropsN);
        //                runN.AppendChild(new Text($"{req.Status.ToString().ToUpper()}    |"));

        //                Run runN1 = paraN.AppendChild(new Run());
        //                RunProperties runPropsN1 = new RunProperties();
        //                runPropsN1.Append(new Color() { Val = "#000000" });
        //                runN1.Append(runPropsN1);
        //                runN1.AppendChild(new Text($"    {req.Specification.Name}    [{passed} passed from {req.ApplicableResults.Count}]    -"));

        //                Run runN2 = paraN.AppendChild(new Run());
        //                RunProperties runPropsN2 = new RunProperties();
        //                runPropsN2.Append(new Color() { Val = "#3381d3" });
        //                runN2.Append(runPropsN2);
        //                runN2.AppendChild(new Text($"    {req.Specification.Cardinality.Description} Requirement"));


        //            }
        //            else
        //            {

        //                Paragraph paraN1 = body.AppendChild(new Paragraph());
        //                Run runN1 = paraN1.AppendChild(new Run());
        //                RunProperties runPropsN1 = new RunProperties();

        //                fontSize = new FontSize() { Val = "30" };
        //                runPropsN1.Append(fontSize);
        //                bold = new Bold(); // Equivale a <w:b/>
        //                runPropsN1.Append(bold);
        //                runPropsN1.Append(new Color() { Val = "#287233" });
        //                runN1.Append(runPropsN1);

        //                runN1.AppendChild(new Text($"{req.Status.ToString().ToUpper()}    |"));


        //                Run runN2 = paraN1.AppendChild(new Run());
        //                RunProperties runPropsN2 = new RunProperties();
        //                runPropsN2.Append(new Color() { Val = "#000000" });
        //                runN2.Append(runPropsN2);
        //                runN2.AppendChild(new Text($"    {req.Specification.Name}    [{passed} passed from {req.ApplicableResults.Count}]    -"));

        //                Run runN3 = paraN1.AppendChild(new Run());
        //                RunProperties runPropsN3 = new RunProperties();
        //                runPropsN3.Append(new Color() { Val = "#0076ad" });
        //                runN3.Append(runPropsN3);
        //                runN3.AppendChild(new Text($"    {req.Specification.Cardinality.Description} Requirement"));



        //            }

        //            foreach (var itm in req.ApplicableResults)
        //            {
        //                if (req.Status == ValidationStatus.Error)
        //                {

        //                    Paragraph paraN1 = body.AppendChild(new Paragraph());
        //                    Run runN1 = paraN1.AppendChild(new Run());
        //                    RunProperties runPropsN1 = new RunProperties();

        //                    //fontSize = new FontSize() { Val = "30" };
        //                    //runPropsN1.Append(fontSize);
        //                    bold = new Bold(); // Equivale a <w:b/>
        //                    runPropsN1.Append(bold);
        //                    runPropsN1.Append(new Color() { Val = "#FF0000" });
        //                    runN1.Append(runPropsN1);

        //                    runN1.AppendChild(new Text($"{StatusIcon(itm.ValidationStatus)} |"));

        //                    foreach (var msg in itm.Messages.Where(m => m.Status != ValidationStatus.Pass))
        //                    {

        //                        Run runN2 = paraN1.AppendChild(new Run());
        //                        RunProperties runPropsN2 = new RunProperties();
        //                        runPropsN2.Append(new Color() { Val = "#000000" });
        //                        runN2.Append(runPropsN2);
        //                        runN2.AppendChild(new Text($"    {req.Specification.Name}    [{passed} passed from {req.ApplicableResults.Count}]    -"));


        //                        Run runN3 = paraN1.AppendChild(new Run());
        //                        RunProperties runPropsN3 = new RunProperties();
        //                        runPropsN3.Append(new Color() { Val = "#0076ad" });
        //                        runN3.Append(runPropsN3);
        //                        runN3.AppendChild(new Text($": {msg?.Reason}\n"));


        //                        Run runN4 = paraN1.AppendChild(new Run());
        //                        RunProperties runPropsN4 = new RunProperties();
        //                        runPropsN4.Append(new Color() { Val = "#0076ad" });
        //                        runN4.Append(runPropsN4);
        //                        runN4.AppendChild(new Text($"     {msg}\n"));


        //                    }

        //                }
        //                else if (req.IsFailure(itm))
        //                {

        //                    Paragraph paraN1 = body.AppendChild(new Paragraph());
        //                    Run runN1 = paraN1.AppendChild(new Run());
        //                    RunProperties runPropsN1 = new RunProperties();

        //                    //fontSize = new FontSize() { Val = "30" };
        //                    //runPropsN1.Append(fontSize);
        //                    bold = new Bold(); // Equivale a <w:b/>
        //                    runPropsN1.Append(bold);
        //                    runPropsN1.Append(new Color() { Val = "#FF0000" });
        //                    runN1.Append(runPropsN1);

        //                    runN1.AppendChild(new Text($"{StatusIcon(itm.ValidationStatus)} |"));


        //                    Run runN2 = paraN1.AppendChild(new Run());
        //                    RunProperties runPropsN2 = new RunProperties();
        //                    runPropsN2.Append(new Color() { Val = "#000000" });
        //                    runN2.Append(runPropsN2);
        //                    runN2.AppendChild(new Text($"{itm.Requirement?.Name} {itm.Requirement?.Description} {itm.FullEntity}\n"));




        //                    foreach (var msg in itm.Messages.Where(m => m.Status != ValidationStatus.Pass))
        //                    {
        //                        var msgtxt = msg.ToString()
        //                            //.Replace("{", "{{")
        //                            //    .Replace("}", "}}")
        //                            ;

        //                        ;
        //                        Run runN3 = paraN1.AppendChild(new Run());
        //                        RunProperties runPropsN3 = new RunProperties();
        //                        runPropsN3.Append(new Color() { Val = "#FF0000" });
        //                        runN3.Append(runPropsN3);
        //                        runN3.AppendChild(new Text($"     {msgtxt}\n"));


        //                    }
        //                }
        //                //else
        //                //{
        //                //    //run2.AppendChild(new Text($"  {StatusIcon(itm.ValidationStatus)} {itm.Requirement?.Name} {itm.Requirement?.Description}"));

        //                //    //console?.WriteDetail(console.GetColorForStatus(itm.ValidationStatus), "  " + StatusIcon(itm.ValidationStatus))
        //                //    //    .WriteDetail(DarkGray, $"{itm.Requirement?.Name} {itm.Requirement?.Description}")
        //                //    //    .WriteDetail(ConsoleColor.Gray, $"{itm.FullEntity}\n");
        //                //    //foreach (var msg in itm.Messages.Where(m => m.Status == ValidationStatus.Pass))
        //                //    //{
        //                //    //    console?.WriteTrace(DarkGray, $"     {msg}\n");
        //                //    //}
        //                //}
        //                //Console.Write(".");
        //            }
        //            //run2.AppendChild(new Text("\n"));

        //        }



        //        //console?.WriteDetailLine(ConsoleColor.Blue, $"  🔎  For {req.Specification.Applicability.GetApplicabilityDescription().SplitClauses()}\n");
        //        //if (req.Specification.Cardinality.AllowsRequirements)
        //        //    console?.WriteDetailLine(DarkGreen, $"  📏  It is {req.Specification.Cardinality.Description} that elements {req.Specification.Requirement?.GetRequirementDescription().SplitClauses()}\n");

        //        //run2.AppendChild(new Text("PLUTO"));

        //        // Salva il documento.
        //        mainPart.Document.Save();
        //    }
        //}

#if SQLite
    private static IModel BuildModelSqlLite(string ifcFile)
    {
        if(!File.Exists(ifcFile))
        {
            throw new FileNotFoundException(ifcFile);
        }
        using (var ifcStream = File.Open(ifcFile, FileMode.Open))
        {
            var file = Path.ChangeExtension(ifcFile, "db");
            var flexDb = new IfcFlexDb(file);

            flexDb.Open(file);
            flexDb.ImportStep21(ifcStream);

            return flexDb;
        }
    }

    private static void OptimiseActivationStrategy(IfcFlexDb model)
    {
        model.AddActivationDepth<IIfcRelDefinesByProperties>(2);
        
        model.AddActivationDepth<IIfcRelDefinesByType>(2);
        model.AddActivationDepth<IIfcRelAssociatesClassification>(2);
        model.AddActivationDepth<IIfcRelAggregates>(2);
    }
#endif

    }
}
