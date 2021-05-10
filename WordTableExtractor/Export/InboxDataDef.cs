using System;
using System.Collections.Generic;
using System.Text;

namespace WordTableExtractor.Export
{
    public class InboxDataDef
    {
        public int ID { get; set; }
        public string Type { get; set; }

        public string Stakeholder { get; set; }
        public string BusinessValue { get; set; }
        public string EVPrioritization { get; set; }
        public string Summary { get; set; }
        public string Safestate { get; set; }
        public string ASIL { get; set; }
        public string FTT { get; set; }
        public string Component { get; set; }
        public string Status { get; set; }
        public string SubmittedBy { get; set; }
        public string SubmittedAt { get; set; }
        public string ModifiedAt { get; set; }
        public string AssignedTo { get; set; }
        public string Team { get; set; }
        public string LeadTeam { get; set; }
        public string MinervaTeam { get; set; }
        public string MinervaLeadTeam { get; set; }
        public string BesondereMerkmale { get; set; }
        public string Objekttyp { get; set; }
        public string StatusProductOfferAndSpecification { get; set; }
        public string StatusTechnicalProductDesign { get; set; }
        public string StatusSensors { get; set; }
        public string StatusIPNextHWComponent { get; set; }
        public string StatusIPNextSubSystems { get; set; }
        public string StatusBKOverarching { get; set; }
        public string StatusPlattformSW { get; set; }
        public string StatusVerificationAndValidation { get; set; }
        public string Description { get; set; }
    }
}
