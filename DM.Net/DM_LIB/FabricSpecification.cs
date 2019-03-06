/*
 * C#
 * User: CRuff
 * Date: 3/6/2019
 * Time: 8:35 AM
 * 
 * 
 */
using System;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace DM_Lib
{
	/// <summary>
	/// Description of FabricSpecification.
	/// This specification is based on a style spec
	/// but adds additional parameters.
	/// </summary>
	public class FabricSpecification : StyleSpecification
	{
		public string VisualCharacteristics { get; set; }
		public float Width { get; set; }
		public float FringeLength { get; set; }	
		public float Thickness { get; set; }
		public string LenoType { get; set; }
		public float CoreID { get; set; }
		public float CoreOD { get; set; }
		public float CoreLength { get; set; }
		public float RollLength { get; set; }
		public float PalletSize { get; set; }
		public string PackagingType { get; set; }
		public string InspectionType { get; set; }
		public bool CustomerFurnished { get; set; }		
		public float ShelfLife { get; set; }
		public string StorageConditions { get; set; }
		public string QualityClausesOnPO { get; set; }
		
		public override string SpecType
        {
            get
            {
                return "fabric";
            }
        }
		
		public FabricSpecification()
		{
			// Default constructor
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
            builder.AppendFormat("Visual Characteristics : {0}\n", VisualCharacteristics);
            builder.AppendFormat("Width : {0}\n", Width);
            builder.AppendFormat("FringeLength : {0}\n", FringeLength);
            builder.AppendFormat("Thickness : {0}\n", Thickness);
            builder.AppendFormat("Core Inner Diameter : {0}\n", CoreID);
            builder.AppendFormat("Core Outter Diameter : {0}\n", CoreOD);
            builder.AppendFormat("Leno : {0}\n", LenoType);
            builder.AppendFormat("Core Length : {0}\n", CoreLength);
            builder.AppendFormat("Pallet Size : {0}\n", PalletSize);
            builder.AppendFormat("Packaging : {0}\n", PackagingType);
            builder.AppendFormat("Inspection : {0}\n", InspectionType);
            builder.AppendFormat("Customer Furnished : {0}\n", CustomerFurnished);
            builder.AppendFormat("Shelf Life : {0}\n", ShelfLife);
            builder.AppendFormat("Storage Conditions : {0}\n", StorageConditions);
            builder.AppendFormat("Quality Clauses as referenced on Purchase Order : {0}\n", QualityClausesOnPO);
            
            return builder.ToString();
        }
		
	}
}
