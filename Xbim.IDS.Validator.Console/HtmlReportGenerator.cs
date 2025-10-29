// HtmlReportGenerator.cs

using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.CommandLine;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using Xbim.IDS.Validator.Common;
using Xbim.IDS.Validator.Core;
using Xbim.Ifc4.Interfaces;

namespace IdsReportGenerator
{
    public class HtmlReportGenerator
    {
        // Limita il numero di righe da mostrare nelle tabelle per non renderle troppo lunghe
        private const int MaxRowsToShowInTable = 20;
        private const int HeadRows = 10;
        private const int TailRows = 5;

        public void Generate(ReportData data, string outputPath)
        {
            var sb = new StringBuilder();

            WriteHeader(sb, data.ProjectName);
            WriteBody(sb, data);
            WriteFooter(sb);

            File.WriteAllText(outputPath, sb.ToString());

        }

        

        private void WriteHeader(StringBuilder sb, string title)
        {
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html lang=\"it\">");
            sb.AppendLine("<head>");
            sb.AppendLine("    <meta charset=\"utf-8\" />");
            sb.AppendLine("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
            sb.AppendLine($"    <title>{WebUtility.HtmlEncode(title)}</title>");
            sb.AppendLine(GetCssStyles()); // Includi il CSS
            sb.AppendLine("</head>");
        }

        private void WriteBody(StringBuilder sb, ReportData data)
        {
            sb.AppendLine("<body>");

            // Intestazione del report
            sb.AppendLine("    <header>");
            sb.AppendLine($"        <h1>{WebUtility.HtmlEncode(data.ProjectName)}</h1>");
            sb.AppendLine($"        <p><strong>{data.Timestamp:yyyy-MM-dd HH:mm:ss}</strong></p>");
            sb.AppendLine("    </header>");

            // Riepilogo generale
            WriteOverallSummary(sb, data.Summary);

            sb.AppendLine("    <hr>");

            // Scorre ogni specifica e crea una sezione
            foreach (var specResult in data.SpecificationResults)
            {
                WriteSpecificationSection(sb, specResult);
            }

            sb.AppendLine("</body>");
        }

        private void WriteOverallSummary(StringBuilder sb, OverallSummaryData summary)
        {
            sb.AppendLine("    <h2>Summary</h2>");
            sb.AppendLine("    <div class=\"container\">");
            string statusClass = summary.Status.ToString().ToLower();
            sb.AppendLine($"        <div class=\"{statusClass} percent\" style=\"width: {summary.Percentage}%;\">{summary.Percentage}%</div>");
            sb.AppendLine("    </div>");
            sb.AppendLine("    <p>");
            sb.AppendLine($"        <span class=\"item {statusClass}\">{summary.Status}</span>");
            sb.AppendLine($"        <span class=\"item\">");
            sb.AppendLine($"            Specifications passed: <strong>{summary.SpecsPassed}</strong> / <strong>{summary.TotalSpecs}</strong>");
            sb.AppendLine("        </span>");
            sb.AppendLine($"        <span class=\"item\">");
            sb.AppendLine($"            Requirements passed: <strong>{summary.ReqsPassed}</strong> / <strong>{summary.TotalReqs}</strong>");
            sb.AppendLine("        </span>");
            sb.AppendLine($"        <span class=\"item\">");
            sb.AppendLine($"            Checks passed: <strong>{summary.ChecksPassed}</strong> / <strong>{summary.TotalChecks}</strong>");
            sb.AppendLine("        </span>");
            sb.AppendLine("    </p>");
        }

        private void WriteSpecificationSection(StringBuilder sb, SpecificationResultData spec)
        {
            sb.AppendLine("    <section style=\"clear: both; overflow: hidden;\">");
            string statusClass = spec.Status.ToString().ToLower();

            // Colonna sinistra con le info
            sb.AppendLine("        <div class=\"info\">");
            sb.AppendLine($"            <h2>{WebUtility.HtmlEncode(spec.Name)}</h2>");
            sb.AppendLine($"            <p>{WebUtility.HtmlEncode(spec.Description)}</p>");
            sb.AppendLine("            <div class=\"container\">");
            sb.AppendLine($"                <div class=\"{statusClass} percent\" style=\"width: {spec.Percentage}%;\">{spec.Percentage}%</div>");
            sb.AppendLine("            </div>");
            sb.AppendLine("            <p>");
            sb.AppendLine($"                <span class=\"item {statusClass}\">{spec.Status}</span>");
            sb.AppendLine($"                <span class=\"item\">");
            sb.AppendLine($"                    Checks passed: <strong>{spec.ChecksPassed}</strong> / <strong>{spec.TotalChecks}</strong>");
            sb.AppendLine("                </span>");
            sb.AppendLine($"                <span class=\"item\">");
            sb.AppendLine($"                    Elements passed: <strong>{spec.ElementsPassed}</strong> / <strong>{spec.TotalElements}</strong>");
            sb.AppendLine("                </span>");
            sb.AppendLine("            </p>");
            sb.AppendLine("        </div>");

            // Colonna destra con i risultati
            sb.AppendLine("        <div class=\"results\">");
            sb.AppendLine("            <p><strong>Applicability</strong></p>");
            sb.AppendLine("            <ul>");
            sb.AppendLine($"                <li>{WebUtility.HtmlEncode(spec.Applicability)}</li>");
            sb.AppendLine("            </ul>");
            sb.AppendLine("            <p><strong>Requirements</strong></p>");
            sb.AppendLine("            <ol>");

            foreach (var req in spec.Requirements)
            {
                WriteRequirementItem(sb, req);
            }

            sb.AppendLine("            </ol>");
            sb.AppendLine("        </div>");

            sb.AppendLine("    </section>");
        }

        private void WriteRequirementItem(StringBuilder sb, RequirementResultData req)
        {
            string statusClass = req.Status.ToString().ToLower();
            sb.AppendLine($"                <li class=\"{statusClass}\">");
            sb.AppendLine("                    <details>");
            sb.AppendLine("                        <summary>");
            sb.AppendLine($"                            {WebUtility.HtmlEncode(req.Description)}");
            sb.AppendLine("                        </summary>");

            // Tabella degli elementi passati (se ce ne sono)
            if (req.PassingElements.Any())
            {
                WriteElementTable(sb, req.PassingElements, true);
            }

            // Tabella degli elementi falliti (se ce ne sono)
            if (req.FailingElements.Any())
            {
                WriteElementTable(sb, req.FailingElements, false);
            }

            sb.AppendLine("                    </details>");
            sb.AppendLine("                </li>");
        }

        private void WriteElementTable(StringBuilder sb, List<ElementInfo> elements, bool isPassing)
        {
            string statusClass = isPassing ? "pass" : "fail";
            sb.AppendLine($"                        <table class=\"{statusClass}\">");
            sb.AppendLine("                            <thead>");
            sb.AppendLine("                                <tr>");
            sb.AppendLine("                                    <th>Class</th>");
            sb.AppendLine("                                    <th>PredefinedType</th>");
            sb.AppendLine("                                    <th>Name</th>");
            sb.AppendLine("                                    <th>Description</th>");
            if (!isPassing) sb.AppendLine("                                    <th>Warning</th>");
            sb.AppendLine("                                    <th>GlobalId</th>");
            sb.AppendLine("                                    <th>Tag</th>");
            sb.AppendLine("                                </tr>");
            sb.AppendLine("                            </thead>");
            sb.AppendLine("                            <tbody>");

            if (elements.Count <= MaxRowsToShowInTable)
            {
                foreach (var el in elements)
                {
                    WriteElementRow(sb, el, isPassing);
                }
            }
            else // Logica per abbreviare tabelle lunghe
            {
                // Mostra le prime N righe
                for (int i = 0; i < HeadRows; i++)
                {
                    WriteElementRow(sb, elements[i], isPassing);
                }
                                // Riga "... more items"
                long hiddenCount = elements.Count - (HeadRows + TailRows);
                string elType = WebUtility.HtmlEncode(elements.First().Name.Split(':').FirstOrDefault());
                sb.AppendLine("                                <tr>");
                sb.AppendLine($"                                    <td colspan=\"7\">... {hiddenCount} more of the same element type ({elType}) not shown ...</td>");
                sb.AppendLine("                                </tr>");

                // Mostra le ultime M righe
                for (int i = elements.Count - TailRows; i < elements.Count; i++)
                {
                    WriteElementRow(sb, elements[i], isPassing);
                }
            }

            sb.AppendLine("                            </tbody>");
            sb.AppendLine("                        </table>");
        }

        private void WriteElementRow(StringBuilder sb, ElementInfo el, bool isPassing)
        {
            sb.AppendLine("                                <tr>");
            //sb.AppendLine($"                                    <td>{WebUtility.HtmlEncode(el.FullEntity)}</td>");
            //sb.AppendLine($"                                    <td>{WebUtility.HtmlEncode(el.PredefinedType ?? "None")}</td>");
            //sb.AppendLine($"                                    <td>{WebUtility.HtmlEncode(el.Name)}</td>");
            //sb.AppendLine($"                                    <td>{WebUtility.HtmlEncode(el.Description ?? "None")}</td>");
            //if (!isPassing) sb.AppendLine($"                                    <td>{WebUtility.HtmlEncode(el.Warning)}</td>");
            //sb.AppendLine($"                                    <td>{WebUtility.HtmlEncode(el.GlobalId)}</td>");
            //sb.AppendLine($"                                    <td>{WebUtility.HtmlEncode(el.Tag)}</td>");
            sb.AppendLine("                                </tr>");
        }

        private void WriteFooter(StringBuilder sb)
        {
            sb.AppendLine("    <hr>");
            sb.AppendLine("    <footer>");
            sb.AppendLine("        <p>");
            //sb.AppendLine("            Report by the <a href=\"https://bonsaibim.org/\">Bonsai</a> and <a href=\"http://ifcopenshell.org/\">IfcOpenShell</a>.");
            sb.AppendLine("        </p>");
            sb.AppendLine("    </footer>");
            sb.AppendLine("</html>");
        }

        private string GetCssStyles()
        {
            // Copia e incolla l'intero blocco <style> dal file HTML di esempio.
            return @"
    <style>
        :root {
            --green: #97cc64;
            --light-green: #b6cca1;
            --red: #fb5a3e;
            --light-red: #fbb4a8;
            --skipped-bg: #8b8d8f;
            --skipped-light-bg: #f5f5f5;
        }
        body { font-family: 'Arial', sans-serif; padding: 10px 40px; }
        section { padding: 15px; border-radius: 5px; border: 1px solid #eee; margin-bottom: 15px; }
        section>h2 { margin-top: 0px; }
        span.time { color: #999; font-style: italic; float: right; }
        span.step-time { float: right; color: #555; font-size: 0.8em; font-style: italic; }
        span.item { padding: 5px; border-radius: 5px; margin-right: 5px; border: 1px solid #eee; display: inline-block; margin-bottom: 5px;}
        span.item.pass, span.item.fail, span.item.skipped { color: #FFF; font-weight: bold; border: 0px; }
        p.fail { background-color: #fb5a3e; padding: 5px; border-radius: 5px; color: #fff; }
        p.unspecified { background-color: #994f00; padding: 5px; border-radius: 5px; color: #fff; }
        p.skipped { background-color: #8b8d8f; padding: 5px; border-radius: 5px; color: #fff; }
        p.description { background-color: #eee; border-radius: 5px; padding: 20px; margin-left: auto; margin-right: auto; display: inline-block; font-weight: bold;}
        li { padding: 10px; font-family: monospace; border-radius: 5px; margin-bottom: 5px; list-style: none;}
        li.pass { background-color: var(--light-green); color: #333; }
        li.fail { background-color: var(--light-red); color: #900; }
        li.unspecified { background-color: #ffd37f; color: #a30; }
        li.skipped { background-color: var(--skipped-light-bg); color: #333; }
        li p { margin-bottom: 0px; }
        footer { color: #999; font-size: 0.8em; }
        header { text-align: center; }
        hr { margin: 20px; margin-left: 0px; margin-right: 0px;  border: none; border-top: 1px solid #ccc; }
        summary { display: flex; cursor: pointer; }
        summary::-webkit-details-marker { display: none; }
        * {box-sizing:border-box}
        .container { width: 100%; background-color: #ddd; border-radius: 5px}
        .percent { text-align: left; padding-top: 5px; padding-left: 5px; padding-bottom: 5px; color: white; border-radius: 5px; white-space: nowrap; font-weight: bold;}
        .pass { background-color: var(--green); }
        .fail { background-color: var(--red); }
        .skipped { background-color: var(--skipped-bg); }
        table { width: 100%; border-spacing: 0; border-radius: 5px; margin-top: 10px; }
        th, td { padding: 5px; color: #000; text-align: left; }
        table.pass { border-bottom: 2px solid var(--green); }
        table.fail { border-bottom: 2px solid var(--red); }
        table thead>tr { font-weight: bold; color: #000; }
        table.pass tbody tr { border-bottom: 1px solid var(--green); }
        table.fail tbody tr { border-bottom: 1px solid var(--red); }
        tbody tr:nth-child(odd) { background-color: rgba(1, 1, 1, 0.05); }
        tbody tr:nth-child(even) { background-color: rgba(1, 1, 1, 0.1); }
        tbody tr:hover { background-color: rgba(0, 0, 0, 0); }
        div.info { float: left; width: 30%; padding-right: 20px;}
        div.results { float: left; width: calc(70% - 20px); margin-left: 20px; padding: 20px; border-radius: 5px; background-color: #fafafa; }
        ol { padding-left: 0; }
    </style>";
        }

       
    }

}