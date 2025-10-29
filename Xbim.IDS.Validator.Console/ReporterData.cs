// ReportData.cs
using IdsReportGenerator;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using Xbim.IDS.Validator.Core;

namespace IdsReportGenerator
{
    // Enum per lo stato del risultato
    public enum ResultStatus { Pass, Fail, Skipped }

    // Dati principali del report
    public class ReportData
    {
        public string ProjectName { get; set; }
        public string Timestamp { get; set; }
        public OverallSummaryData Summary { get; set; }
        public List<SpecificationResultData> SpecificationResults { get; set; } = new List<SpecificationResultData>();
    }

    // Riepilogo generale
    public class OverallSummaryData
    {
        public ResultStatus Status { get; set; }
        public int Percentage { get; set; }
        public int SpecsPassed { get; set; }
        public int TotalSpecs { get; set; }
        public int ReqsPassed { get; set; }
        public int TotalReqs { get; set; }
        public long ChecksPassed { get; set; }
        public long TotalChecks { get; set; }
    }

    // Risultato di una singola specifica (es. "Pset_5D")
    public class SpecificationResultData
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public ResultStatus Status { get; set; }
        public int Percentage { get; set; }
        public long ChecksPassed { get; set; }
        public long TotalChecks { get; set; }
        public long ElementsPassed { get; set; }
        public long TotalElements { get; set; }
        public string Applicability { get; set; }
        public List<RequirementResultData> Requirements { get; set; } = new List<RequirementResultData>();
    }

    // Risultato di un singolo requisito
    public class RequirementResultData
    {
        public string Description { get; set; }
        public ResultStatus Status { get; set; }
        public List<ElementInfo> PassingElements { get; set; } = new List<ElementInfo>();
        public List<ElementInfo> FailingElements { get; set; } = new List<ElementInfo>();
    }

    // Informazioni su un singolo elemento IFC
    //public class ElementInfo
    //{
    //    public string FullEntity { get; set; }
    //    public string Warning { get; set; } // Solo per elementi falliti
}
public class ElementInfo
{
    public string IfcClass { get; set; }
    public string PredefinedType { get; set; }
    public string Name { get; set; }
    public string Description { get; set; }
    public string GlobalId { get; set; }
    public string Tag { get; set; }
    public string Warning { get; set; } // Solo per elementi falliti
}
public class strRigaExcel
{
    public string Requirement { get; set; }
    public string Problem { get; set; }
    public ValidationMessage Messag { get; set; }
    public string Classe { get; set; }

    public string Nome { get; set; }

    public string GlobalID { get; set; }
    public string NumTag { get; set; }
    public IdsValidationResult Elemento { get; set; }
}