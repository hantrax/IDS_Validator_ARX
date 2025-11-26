using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using Xbim.BCF;
using Xbim.BCF.XMLNodes;
using Xbim.Common;
using Xbim.IDS.Validator.Common;
using Xbim.IDS.Validator.Console.Internal;
using Xbim.IDS.Validator.Core;
using Xbim.IDS.Validator.Core.Interfaces;
using Xbim.Ifc4.Interfaces;
using Xbim.InformationSpecifications;
using Xbim.IO.CobieExpress;
using Xbim.IO.Memory;
using BcfToolkit.Builder.Bcf30;
using ARXCommand;

using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System.CommandLine.Invocation;
using System.Diagnostics;

using static System.ConsoleColor;
using static Xbim.IDS.Validator.Console.CliOptions;






namespace ARXCommand
{
    public class ARXNewCommand
    {
        //public void BCFExport(ValidationOutcome outcome, string pathModello, string bcfPath)
        //{
        //    try
        //        {

        //        var bcf = new BCF();


        //        foreach (ValidationRequirement req in outcome.ExecutedRequirements)
        //        {
                    
        //            var snapshotData = System.IO.File.ReadAllBytes("snapshot0.png");

        //            var snapshotData1 = BcfToolkit.Builder.Bcf30.Base64FileData.FromFile("snapshot0.png");



        //            var builder = new BcfBuilder();
        //            var bcf1 = builder
        //              .AddMarkup(m => m
        //                .SetTitle("Simple title")
        //                .SetDescription("This is a description")
        //                .AddLabel("Architecture")
        //                .SetPriority("Critical")
        //                .SetTopicType("Clash")
        //                .SetTopicStatus("Active")
        //                .AddComment(c => c
        //                  .SetComment("This is a comment")
        //                  .SetDate(DateTime.Now)
        //                  .SetAuthor("jimmy@page.com"))
        //                .AddViewPoint(v => v
        //                    .SetVisualizationInfo(vis => vis
        //                        .SetPerspectiveCamera(cam => cam
        //                            .SetCamera(camera => camera
        //                                .SetViewPoint(10, 10, 10)
        //                            )
        //                        )
        //                    )
        //                    .SetSnapshotData(new BcfToolkit.Model.FileData
        //                    {
        //                        Data = snapshotData,
        //                        FileName = "snapshot0.png",
        //                        MimeType = "image/png"
        //                    })
        //                  )
        //               .SetExtensions(e => e
        //                .AddPriority("Critical")
        //                .AddPriority("Major")
        //                .AddPriority("Normal")
        //                .AddPriority("Minor")
        //                .AddTopicType("Issue")
        //                .AddTopicType("Fault")
        //                .AddTopicType("Clash")
        //                .AddTopicType("Remark")
        //                .AddTopicLabel("Architecture")
        //                .AddTopicLabel("Structure")
        //                .AddTopicLabel("MEP"))
        //              .SetProject(p => p
        //                .SetProjectId("projectId")
        //                .SetProjectName("My project"))
        //              .SetDocumentInfo(dI => dI
        //                .AddDocument(d => d
        //                  .SetFileName("document.pdf")
        //                  .SetDescription("This is a document")))
        //              .Build();



        //            var passed = req.PassedResults.Count();

        //            //run2.AppendChild(new Text(req.Status.ToString()));

        //            if (req.Status == ValidationStatus.Fail || req.Status == ValidationStatus.Error)
        //            {
        //                bcf.Project = new ProjectXMLFile
        //                {
        //                    Project = new BCFProject
        //                    {
        //                        Name = "Prova BCF",
        //                        ProjectId = "9B981279-4B63-46A0-ABF0-432E27B5ADC0"
        //                    }
        //                };
                        
        //                bcf.Version = new VersionXMLFile("1.0");
        //                bcf.Topics = new List<Topic> {
        //            new Topic {
        //                Markup = new MarkupXMLFile {
        //                    Header = new BCFHeader {
        //                        Files = new List<BCFFile> {
        //                            new BCFFile {
        //                                Date = DateTime.Now,
        //                                Filename = Path.GetFileName(pathModello),
        //                                IfcProject = "prova",
        //                                IfcSpatialStructureElement = "9B981279-4B63-46A0-ABF0-432E27B5ADC7",
                                       

        //                            }
        //                        }
        //                    },
                            
                            
        //                    Comments = new List<BCFComment> {
                                
        //                        new BCFComment(
        //                            Guid.NewGuid(),
        //                            Guid.NewGuid(),
        //                            "open",
        //                            DateTime.Now,
        //                            req.Status.ToString().ToUpper(),
        //                            req.Specification.Cardinality.Description) {
        //                            Topic = new AttrIDNode(new Guid("9B981279-4B63-46A0-ABF0-432E27B5ADC0"))
        //                        }
        //                    },







        //                    Topic = new BCFTopic(new Guid("9B981279-4B63-46A0-ABF0-432E27B5ADC0"), req.Specification.Guid),
                            
        //                    Viewpoints = new List<BCFViewpoint> {
        //                        new BCFViewpoint(Guid.NewGuid()) { Snapshot = "snapshot02.png", Viewpoint = "Base Viewpoint"}
        //                    }
        //                },
        //                Snapshots = new List<KeyValuePair<string, byte[]>> {
        //                    new KeyValuePair<string, byte[]>( "snapshot02.png", File.ReadAllBytes("snapshot02.png"))
        //                },
        //                Visualization = new VisualizationXMLFile {

        //                }
        //            }
        //        };
        //            }

        //        }

        //        using (var output = File.Create(bcfPath))
        //        {
        //            var data = bcf.Serialize();
        //            data.CopyTo(output);
        //            output.Close();
        //        }



        //    }
        //    catch (Exception ioExp)
        //    {
        //        Console.WriteLine($"Errore durante l'eliminazione del file BCF esistente: {ioExp.Message}");
        //        return;
        //    }






        //}
        public void SaveExcelFile(ValidationOutcome outcome, string fileExcel)
        {

            try
            {
                using var workbook = new XLWorkbook();
                string strNameEntity = "";
                int Riga = 2;
                int Colonna = 1;
                var tmpRiga = new strRigaExcel();
                var tmpRigaLista = new List<strRigaExcel>();

                foreach (var requirement in outcome.ExecutedRequirements)
                {
                    if (requirement.Status is ValidationStatus.Fail)
                    {
                        var sheetName = SanitizeSheetName(requirement.Specification.Name);
                        var worksheet = workbook.Worksheets.Add(sheetName);
                        worksheet.Cell(1, 1).Value = "ERRORE";
                        worksheet.Cell(1, 1).Style.Fill.SetBackgroundColor(XLColor.Gray);
                        worksheet.Cell(1, 4).Value = "NOME OGGETTO";
                        worksheet.Cell(1, 5).Value = "GUID";
                        worksheet.Cell(1, 6).Value = "RevitID";
                        worksheet.Cell(1, 7).Value = "Elemento";


                        // Dettagli errori
                        var failedEntities = requirement.ApplicableResults.Where(e => e.ValidationStatus != ValidationStatus.Pass);
                        if (failedEntities.Any())
                        {

                            foreach (var entity in failedEntities)
                            {

                                var ifcEntity = entity.FullEntity;
                                var ifcRoot = ifcEntity as IIfcRoot;



                                tmpRiga.Elemento = entity;
                                if (entity is IIfcObject failingObject1)
                                {
                                    tmpRiga.GlobalID = failingObject1.GlobalId;

                                }

                                foreach (var message in entity.Messages)
                                {
                                    if (message.Status != ValidationStatus.Pass)
                                    {
                                        string[] words = ifcRoot?.Name.ToString().Split(":");

                                        var numWord = words.Count();


                                        worksheet.Cell(Riga, 1).Value = message.ToString();
                                        worksheet.Cell(Riga, 1).Style.Alignment.WrapText=true;

                                        worksheet.Cell(Riga, 4).Value = ifcRoot?.Name.ToString();
                                        worksheet.Cell(Riga, 5).Value = ifcRoot?.GlobalId.ToString();
                                        worksheet.Cell(Riga, 6).Value = words[numWord-1];
                                        worksheet.Cell(Riga, 7).Value = entity.FullEntity.ToString();
                                        worksheet.Row(Riga).AdjustToContents(1);
                                        //worksheet.Row(Riga).ClearHeight();


                                        Riga += 1;

                                    }


                                }


                            }

                        }
                        Riga = 2;
                        worksheet.Columns().AdjustToContents();
                        worksheet.Rows().AdjustToContents(1);
                    }
                }


                workbook.SaveAs(fileExcel);
                workbook.Dispose();
            }
            catch (Exception ex)
            {
                // Cattura altre possibili eccezioni (es. percorso non valido).
                //Console.WriteLine($"Si è verificato un errore imprevisto: {ex.Message}");
                //Console.WriteLine($"ERRORE: {ex.Message}");
                if (ex.InnerException != null)
                {
                    //Console.WriteLine($"Dettaglio: {ex.InnerException.Message}");
                }
                //Console.WriteLine($"StackTrace: {ex.StackTrace}");
            }
        }

        static string SanitizeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Foglio";
            var invalidChars = new[] { '\\', '/', '?', '*', '[', ']' };
            foreach (char c in invalidChars) { name = name.Replace(c, '_'); }
            if (name.Length > 31) { name = name.Substring(0, 31); }
            return name;
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

        public static void SaveResultsToWORD(ValidationOutcome outcome, string outputPath, string modello, string idsFile, string IFCVersion)
        {
            // Crea un nuovo documento Word.
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
            {
                // Aggiungi una parte principale al documento.
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());


                SectionProperties sectionProps = new SectionProperties();


                PageSize pageSize = new PageSize()
                {
                    Width = (UInt32Value)23810U, // Larghezza A3 in Twips (altezza se verticale)
                    Height = (UInt32Value)16836U  // Altezza A3 in Twips (larghezza se verticale)
                };

                pageSize.Orient = PageOrientationValues.Landscape;

                sectionProps.Append(pageSize);


                // Aggiungi la sezione al corpo del documento
                body.Append(sectionProps);





                RunProperties runProperties1 = new RunProperties(); // Definisce le proprietà del testo (es. colore).
                runProperties1.AppendChild(new Color() { Val = "#3381d3" });
                FontSize fontSize = new FontSize() { Val = "40" };
                runProperties1.Append(fontSize);
                Bold bold = new Bold(); // Equivale a <w:b/>
                runProperties1.Append(bold);
                Paragraph para1 = body.AppendChild(new Paragraph());
                Run run1 = para1.AppendChild(new Run());
                run1.AppendChild(runProperties1);

                run1.AppendChild(new Text($"Nome Modello:   {modello}"));

                RunProperties runProperties2 = new RunProperties(); // Definisce le proprietà del testo (es. colore).
                runProperties2.AppendChild(new Color() { Val = "#3381d3" });
                fontSize = new FontSize() { Val = "40" };
                runProperties2.Append(fontSize);
                bold = new Bold(); // Equivale a <w:b/>
                runProperties2.Append(bold);
                Paragraph para2 = body.AppendChild(new Paragraph());
                Run run2 = para2.AppendChild(new Run());
                run2.AppendChild(runProperties2);

                run2.AppendChild(new Text($"Nome File IDS:   {idsFile}"));

                body.AppendChild(new Paragraph(new Run(new Text(""))));


                RunProperties runProperties3 = new RunProperties(); // Definisce le proprietà del testo (es. colore).
                runProperties3.AppendChild(new Color() { Val = "#3381d3" });
                fontSize = new FontSize() { Val = "40" };
                runProperties3.Append(fontSize);
                bold = new Bold(); // Equivale a <w:b/>
                runProperties3.Append(bold);
                Paragraph para3 = body.AppendChild(new Paragraph());
                Run run3 = para3.AppendChild(new Run());
                run3.AppendChild(runProperties3);

                run3.AppendChild(new Text($"Versione IFC: {IFCVersion} - Data Elaborazione: {DateTime.Now:dd-MM-yyyy HH:mm:ss}"));

                body.AppendChild(new Paragraph(new Run(new Text(""))));





                foreach (ValidationRequirement req in outcome.ExecutedRequirements)
                {
                    var passed = req.PassedResults.Count();

                    //run2.AppendChild(new Text(req.Status.ToString()));

                    if (req.Status == ValidationStatus.Fail || req.Status == ValidationStatus.Error)
                    {

                        Paragraph paraN = body.AppendChild(new Paragraph());
                        Run runN = paraN.AppendChild(new Run());
                        RunProperties runPropsN = new RunProperties();
                        // Imposta il colore a nero. Anche se il nero è il colore predefinito, lo si specifica per chiarezza.
                        fontSize = new FontSize() { Val = "30" };
                        runPropsN.Append(fontSize);
                        bold = new Bold(); // Equivale a <w:b/>
                        runPropsN.Append(bold);
                        runPropsN.Append(new Color() { Val = "#FF0000" });
                        runN.Append(runPropsN);
                        runN.AppendChild(new Text($"{req.Status.ToString().ToUpper()}    |"));

                        Run runN1 = paraN.AppendChild(new Run());
                        RunProperties runPropsN1 = new RunProperties();
                        runPropsN1.Append(new Color() { Val = "#000000" });
                        runN1.Append(runPropsN1);
                        runN1.AppendChild(new Text($"    {req.Specification.Name}    [{passed} passed from {req.ApplicableResults.Count}]    -"));

                        Run runN2 = paraN.AppendChild(new Run());
                        RunProperties runPropsN2 = new RunProperties();
                        runPropsN2.Append(new Color() { Val = "#3381d3" });
                        runN2.Append(runPropsN2);
                        runN2.AppendChild(new Text($"    {req.Specification.Cardinality.Description} Requirement"));


                    }
                    else
                    {

                        Paragraph paraN1 = body.AppendChild(new Paragraph());
                        Run runN1 = paraN1.AppendChild(new Run());
                        RunProperties runPropsN1 = new RunProperties();

                        fontSize = new FontSize() { Val = "30" };
                        runPropsN1.Append(fontSize);
                        bold = new Bold(); // Equivale a <w:b/>
                        runPropsN1.Append(bold);
                        runPropsN1.Append(new Color() { Val = "#287233" });
                        runN1.Append(runPropsN1);

                        runN1.AppendChild(new Text($"{req.Status.ToString().ToUpper()}    |"));


                        Run runN2 = paraN1.AppendChild(new Run());
                        RunProperties runPropsN2 = new RunProperties();
                        runPropsN2.Append(new Color() { Val = "#000000" });
                        runN2.Append(runPropsN2);
                        runN2.AppendChild(new Text($"    {req.Specification.Name}    [{passed} passed from {req.ApplicableResults.Count}]    -"));

                        Run runN3 = paraN1.AppendChild(new Run());
                        RunProperties runPropsN3 = new RunProperties();
                        runPropsN3.Append(new Color() { Val = "#0076ad" });
                        runN3.Append(runPropsN3);
                        runN3.AppendChild(new Text($"    {req.Specification.Cardinality.Description} Requirement"));



                    }

                    foreach (var itm in req.ApplicableResults)
                    {
                        if (req.Status == ValidationStatus.Error)
                        {

                            Paragraph paraN1 = body.AppendChild(new Paragraph());
                            Run runN1 = paraN1.AppendChild(new Run());
                            RunProperties runPropsN1 = new RunProperties();

                            //fontSize = new FontSize() { Val = "30" };
                            //runPropsN1.Append(fontSize);
                            bold = new Bold(); // Equivale a <w:b/>
                            runPropsN1.Append(bold);
                            runPropsN1.Append(new Color() { Val = "#FF0000" });
                            runN1.Append(runPropsN1);

                            runN1.AppendChild(new Text($"{StatusIcon(itm.ValidationStatus)} |"));

                            foreach (var msg in itm.Messages.Where(m => m.Status != ValidationStatus.Pass))
                            {

                                Run runN2 = paraN1.AppendChild(new Run());
                                RunProperties runPropsN2 = new RunProperties();
                                runPropsN2.Append(new Color() { Val = "#000000" });
                                runN2.Append(runPropsN2);
                                runN2.AppendChild(new Text($"    {req.Specification.Name}    [{passed} passed from {req.ApplicableResults.Count}]    -"));


                                Run runN3 = paraN1.AppendChild(new Run());
                                RunProperties runPropsN3 = new RunProperties();
                                runPropsN3.Append(new Color() { Val = "#0076ad" });
                                runN3.Append(runPropsN3);
                                runN3.AppendChild(new Text($": {msg?.Reason}\n"));


                                Run runN4 = paraN1.AppendChild(new Run());
                                RunProperties runPropsN4 = new RunProperties();
                                runPropsN4.Append(new Color() { Val = "#0076ad" });
                                runN4.Append(runPropsN4);
                                runN4.AppendChild(new Text($"     {msg}\n"));


                            }

                        }
                        else if (req.IsFailure(itm))
                        {

                            Paragraph paraN1 = body.AppendChild(new Paragraph());
                            Run runN1 = paraN1.AppendChild(new Run());
                            RunProperties runPropsN1 = new RunProperties();

                            //fontSize = new FontSize() { Val = "30" };
                            //runPropsN1.Append(fontSize);
                            bold = new Bold(); // Equivale a <w:b/>
                            runPropsN1.Append(bold);
                            runPropsN1.Append(new Color() { Val = "#FF0000" });
                            runN1.Append(runPropsN1);

                            runN1.AppendChild(new Text($"{StatusIcon(itm.ValidationStatus)} |"));


                            Run runN2 = paraN1.AppendChild(new Run());
                            RunProperties runPropsN2 = new RunProperties();
                            runPropsN2.Append(new Color() { Val = "#000000" });
                            runN2.Append(runPropsN2);
                            runN2.AppendChild(new Text($"{itm.Requirement?.Name} {itm.Requirement?.Description} {itm.FullEntity}\n"));




                            foreach (var msg in itm.Messages.Where(m => m.Status != ValidationStatus.Pass))
                            {
                                var msgtxt = msg.ToString()
                                    //.Replace("{", "{{")
                                    //    .Replace("}", "}}")
                                    ;

                                ;
                                Run runN3 = paraN1.AppendChild(new Run());
                                RunProperties runPropsN3 = new RunProperties();
                                runPropsN3.Append(new Color() { Val = "#FF0000" });
                                runN3.Append(runPropsN3);
                                runN3.AppendChild(new Text($"     {msgtxt}\n"));


                            }
                        }
                        //else
                        //{
                        //    //run2.AppendChild(new Text($"  {StatusIcon(itm.ValidationStatus)} {itm.Requirement?.Name} {itm.Requirement?.Description}"));

                        //    //console?.WriteDetail(console.GetColorForStatus(itm.ValidationStatus), "  " + StatusIcon(itm.ValidationStatus))
                        //    //    .WriteDetail(DarkGray, $"{itm.Requirement?.Name} {itm.Requirement?.Description}")
                        //    //    .WriteDetail(ConsoleColor.Gray, $"{itm.FullEntity}\n");
                        //    //foreach (var msg in itm.Messages.Where(m => m.Status == ValidationStatus.Pass))
                        //    //{
                        //    //    console?.WriteTrace(DarkGray, $"     {msg}\n");
                        //    //}
                        //}
                        //Console.Write(".");
                    }
                    //run2.AppendChild(new Text("\n"));

                }



                //console?.WriteDetailLine(ConsoleColor.Blue, $"  🔎  For {req.Specification.Applicability.GetApplicabilityDescription().SplitClauses()}\n");
                //if (req.Specification.Cardinality.AllowsRequirements)
                //    console?.WriteDetailLine(DarkGreen, $"  📏  It is {req.Specification.Cardinality.Description} that elements {req.Specification.Requirement?.GetRequirementDescription().SplitClauses()}\n");

                //run2.AppendChild(new Text("PLUTO"));

                // Salva il documento.
                mainPart.Document.Save();
            }
        }

    }
}