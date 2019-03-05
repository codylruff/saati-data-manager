/*
 * C#
 * User: CRuff
 * Date: 2/14/2019
 * Time: 11:11 AM
 * 
 * 
 */
using System;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace DM_Lib
{
    public class StyleSpecification : ISpec
    {
        public float Dtex { get; set; }
        public string Style { get; set; }
        public string WeaveType { get; set; }
        public string YarnType { get; set; }
        public float Denier { get; set; }
        public float MeanWarpCount { get; set; }
        public float MinWarpCount { get; set; }
        public float MaxWarpCount { get; set; }
        public float MeanFillCount { get; set; }
        public float MinFillCount { get; set; }
        public float MaxFillCount { get; set; }
        public float MeanDryWeight { get; set; }
        public float MinDryWeight { get; set; }
        public float MaxDryWeight { get; set; }
        public float MeanConditionedWeight { get; set; }
        public float MinConditionedWeight { get; set; }
        public float MaxConditionedWeight { get; set; }
        public string YarnFinish { get; set; }
        public string YarnCode { get; set; }
        public float MoistureRegain { get; set; }
        public float Twisting { get; set; }
        public string YarnColor { get; set; }
        public string Notes { get; set; }
        public string YarnMerge { get; set; }
        public DateTime TimeStamp { get; set; }
        public bool IsDefault { get; set; }
        public string MaterialId { get; set; }
        public string Revision { get; set; }
        public string SpecType
        {
            get
            {
                return "style";
            }
        }
        
        [JsonIgnore]
        public ISpec ParentSpec { get; set; }

        public StyleSpecification(string style)
        {
            // There is an element type of specification
            this.ParentSpec = null;
            this.Style = style;
            this.MaterialId = style;
        }
        
        public override string ToString()
        {
            var builder = new StringBuilder();
     		
            builder.AppendFormat("Revision : {0}\n", Revision);
            builder.AppendFormat("Dtex : {0}\n", Dtex);
            builder.AppendFormat("Style : {0}\n", Style);
            builder.AppendFormat("Weave Type : {0}\n", WeaveType);
            builder.AppendFormat("Yarn Type : {0}\n", YarnType);
            builder.AppendFormat("Denier : {0}\n", Denier);
            builder.AppendFormat("Warp Count : {0} ({1} to {2})\n", MeanWarpCount, MinWarpCount, MaxWarpCount);
            builder.AppendFormat("Fill Count : {0} ({1} to {2})\n", MeanFillCount, MinFillCount, MaxFillCount);
            builder.AppendFormat("Dry Weight : {0} ({1} to {2})\n", MeanDryWeight, MinDryWeight, MaxDryWeight);
            builder.AppendFormat("Conditioned Weight : {0} ({1} to {2})\n", MeanConditionedWeight, MinConditionedWeight, MaxConditionedWeight);
            builder.AppendFormat("Yarn Finish : {0}\n", YarnFinish);
            builder.AppendFormat("Yarn Code : {0}\n", YarnCode);
            builder.AppendFormat("Moisture Regain : {0}\n", MoistureRegain);
            builder.AppendFormat("Twisting : {0}\n", Twisting);
            builder.AppendFormat("Yarn Color : {0}\n", YarnColor);
            builder.AppendFormat("Notes : {0}\n", Notes);
            builder.AppendFormat("Merge : {0}\n", YarnMerge);

            return builder.ToString();
        }

        public void SetDefaultProperties()
        {
            // Currently these are not calculated but 
            // simply input into the database as data.
        }
    }
}