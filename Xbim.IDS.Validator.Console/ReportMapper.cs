using IdsReportGenerator;
using Xbim.IDS.Validator.Common;

using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Xbim.Ifc;
using Xbim.IDS.Validator.Core.Interfaces;
using Xbim.IDS.Validator.Core.Extensions;

using Xbim.IDS.Validator.Core;
using Xbim.IO;
using Xbim.InformationSpecifications;
using Microsoft.Isam.Esent.Interop;
using System.Collections.Generic;


public class IdsReportMapper
{
    public ReportData Map(ValidationOutcome outcome, string Modello, string verIFC)
    {

        var reportData = new ReportData();
        
        var globSummary = new OverallSummaryData();
        var SpecPassed = 0;

        var totalPassed = 0;
        var totalFailed = 0;
        var totalTested = 0;

        var tmpGlobalStatus = new ResultStatus();
        var requirementFacet = 0;
        var elemTested = 0;
        var elemPassed = 0;

        var listFail = new List<string>();
        var reqPassed = 0;
        var reqFail = 0;


        // Parte generale
        reportData.ProjectName = "Nome Modello: " + Modello; 
        reportData.Timestamp = "Data Controllo: " + DateTime.Now.ToString() + "  --  Versione IFC: " + verIFC;


        switch (outcome.Status.ToString())
        {
            case "Fail":
                tmpGlobalStatus= ResultStatus.Fail;
                break;
            case "Pass":
                tmpGlobalStatus= ResultStatus.Pass;
                SpecPassed += 1;
                break;
            case "Skipped":
                tmpGlobalStatus= ResultStatus.Skipped;
                break;

            default:
                //default
                break;
        }

        foreach (var ExecReq in outcome.ExecutedRequirements)
        {
            if (ExecReq.Specification.Name is "Duna")
            {
                var X = 0;
            }

            var specData = new SpecificationResultData();


            var entitiesTested = ExecReq.ApplicableResults.Count();
            var entitiesPassed = ExecReq.ApplicableResults.Count(e => e.ValidationStatus == ValidationStatus.Pass);

            SpecPassed = outcome.ExecutedRequirements.Count(e => e.Status == ValidationStatus.Pass);


            totalTested += entitiesTested;
            totalPassed += entitiesPassed;
            totalFailed += (entitiesTested - entitiesPassed);



            int failEntity = ExecReq.FailedResults.Count();
            int passEntity = ExecReq.PassedResults.Count();
           



            int percSpec = (int)Math.Abs((double)((double)passEntity/(double)(passEntity+failEntity)) * 100);
                  
            var tmpSpecStatus = new ResultStatus();

            switch (ExecReq.Status.ToString())
            {
                case "Fail":
                    tmpSpecStatus= ResultStatus.Fail;
                    break;
                case "Pass":
                    tmpSpecStatus= ResultStatus.Pass;
                    //SpecPassed += 1;
                    break;
                case "Skipped":
                    tmpSpecStatus= ResultStatus.Skipped;
                    break;

                default:
                    //default
                    break;
            }

            switch (ExecReq.ApplicableResults.Count()) //Leggo lo Status della specifica
            {
                case 0:
                    tmpSpecStatus= ResultStatus.Skipped;
                    requirementFacet = 0;
                    percSpec = 0;
                                       // Riquadro Specification Generali
                    specData.Name = ExecReq.Specification.Name;
                    specData.Description = ExecReq.Specification.Description;
                    specData.Status = tmpSpecStatus;

                    specData.Percentage = 0;

                    specData.ChecksPassed = 0;
                    specData.TotalChecks =  0;
                    specData.ElementsPassed  = 0;
                    specData.TotalElements = 0;
                    specData.Applicability = ExecReq.Specification.Applicability.GetApplicabilityDescription();
                                    

                    //reportData.SpecificationResults.Add(specData);
                    goto a100;
                    break;
            }
            switch (ExecReq.Status.ToString()) 
            {
                case "Fail":
                    tmpSpecStatus= ResultStatus.Fail;
                    
                    break;
                case "Pass":
                    tmpSpecStatus= ResultStatus.Pass;
                   
                    break;
                case "Skipped":
                    tmpSpecStatus= ResultStatus.Skipped;
                    break;

                default:
                    //default
                    break;
            }

            specData.Name= ExecReq.Specification.Name;


            specData.Description = ExecReq.Specification.Description;
            specData.Status = tmpSpecStatus;
            specData.Percentage = percSpec;
           
            specData.ChecksPassed = passEntity * ExecReq.Specification.Requirement.Facets.Count();
            specData.TotalChecks =  (passEntity+failEntity)*ExecReq.Specification.Requirement.Facets.Count;
            specData.ElementsPassed  = passEntity;
            specData.TotalElements = passEntity+failEntity;
            specData.Applicability = ExecReq.Specification.Applicability.GetApplicabilityDescription();

            elemTested += (passEntity+failEntity)*requirementFacet;
            elemPassed += passEntity;

            //Controllo il colore dei requirement
            foreach (var result in ExecReq.ApplicableResults)
            {
                if (result.ValidationStatus == ValidationStatus.Fail)
                {

                    foreach (var mess in result.Messages)
                    {
                        if (mess.Status == ValidationStatus.Fail)
                        {
                            if (!listFail.Contains(mess.Clause.RequirementDescription))
                            {
                                listFail.Add(mess.Clause.RequirementDescription);
                            };
                        }
                    }


                    //foreach (var requirement in result.Requirement.Facets)
                    //{
                    //    if (!listFail.Contains(requirement.RequirementDescription))
                    //    {
                    //      listFail.Add(requirement.RequirementDescription);
                    //    };
                    //}
                }
            }

            //Setto lo status del controllo
            reqFail= 0;
            reqPassed= 0;

            foreach (var requir in ExecReq.Specification.Requirement.RequirementOptions)
            {
                var reqData = new RequirementResultData();
                reqData.Description=requir.RelatedFacet.RequirementDescription;

                if (listFail.Contains(requir.RelatedFacet.RequirementDescription))
                {
                    reqData.Status = ResultStatus.Fail;
                    reqFail += 1;

                }
                else
                {
                    reqData.Status = ResultStatus.Pass;
                    reqPassed += 1;
                }
                specData.Requirements.Add(reqData);


            };

            //specData.ChecksPassed = reqPassed * specData.ElementsPassed;
            

a100:            reportData.SpecificationResults.Add(specData);

        } //fine ciclo
        
        //Dati Del Sommario
        globSummary.Status = tmpGlobalStatus;
        globSummary.Percentage = (int)Math.Abs((double)((double)totalPassed/(double)(totalTested) * 100));
        globSummary.SpecsPassed = SpecPassed;
        globSummary.TotalSpecs = outcome.ExecutedRequirements.Count;
        globSummary.ReqsPassed = reqPassed;
        globSummary.TotalReqs = reqPassed+reqFail;
        globSummary.ChecksPassed = totalPassed;
        globSummary.TotalChecks = totalTested;
        
      
        reportData.Summary= globSummary;

        return reportData;
    }
}

