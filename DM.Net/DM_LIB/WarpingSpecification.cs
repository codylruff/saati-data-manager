/*
 * C#
 * User: CRuff
 * Date: 2/14/2019
 * Time: 11:11 AM
 * 
 * 
 */
using System;
using System.Text;
using Newtonsoft.Json;

namespace DM_Lib
{
    public class WarpingSpecification : ISpec
    {
        const float HighDtex = 3000;
        public string MaterialNumber { get; set; }
        public string MaterialDescription { get; set; }
        public float FinalWidthCm { get; set; }
        public float EndsPerInch { get; set; }
        public float Dtex { get; set; }
        public float WarpingSpeed { get; set; }
        public float BeamingSpeed { get; set; }
        public float CrossWinding { get; set; }
        public float DentsPerCm { get; set; }
        public float EndsPerDent { get; set; }
        public bool IsSWrapped { get; set; }
        public float NumberOfEnds { get; set; }
        public float BeamWidth { get; set; }
        public float BeamingTension { get; set; }
        public float WarpingTension { get; set; }
        public float K1 { get; set; }
        public float K2 { get; set; }
        public StyleSpecification StyleSpec { get; set; }
        public string Style { get; set; }
        public float WarpDensityStart { get; set; }
        public float WarpDensityMeasured { get; set; }
        public float CompactionFactor { get; set; }
        public bool EvenerRoller { get; set; }
        public bool MeasuringPhase { get; set; }
        public int ReedShape { get; set; }
        public float YarnLeaseFeed { get; set; }
        public string YarnSupplier { get; set; }
        public DateTime TimeStamp { get; set; }
        public bool IsDefault { get; set; }
        public string MaterialId { get; set; }
        public string Revision { get; set; }
        public string SpecType {
            get
            {
                return "warping";   
            }
        }
        
        [JsonIgnore]
        public ISpec ParentSpec { get; set; }

        public WarpingSpecification()
        {
			// Default Constructor
        }

        public WarpingSpecification(string materialId)
        {
            this.MaterialId = materialId;
        }

        public WarpingSpecification(string materialId, StyleSpecification styleSpec)
        {
            this.MaterialId = materialId;
            this.StyleSpec = styleSpec;
            this.ParentSpec = styleSpec;
            this.SetDefaultProperties();
        }

        public override string ToString()
        {
            var builder = new StringBuilder();
			
            builder.AppendFormat("Revision : {0}\n", Revision);
            builder.AppendFormat("Material Number : {0}\n", MaterialNumber);
            builder.AppendFormat("Description : {0}\n", MaterialDescription);
            builder.AppendFormat("Final Width [cm] : {0}\n", FinalWidthCm);
            builder.AppendFormat("End Count : {0}\n", EndsPerInch);
            builder.AppendFormat("Dtex : {0}\n", Dtex);
            builder.AppendFormat("Warping Speed [m/min] : {0}\n", WarpingSpeed);
            builder.AppendFormat("Beaming Speed [m/min] : {0}\n", BeamingSpeed);
            builder.AppendFormat("Cross Winding : {0}\n", CrossWinding);
            builder.AppendFormat("Dents / cm : {0}\n", DentsPerCm);
            builder.AppendFormat("Ends / Dent : {0}\n", EndsPerDent);
            builder.AppendFormat("S Wrap : {0}\n", IsSWrapped ? "ON" : "OFF");
            builder.AppendFormat("Total Number Of Ends : {0}\n", NumberOfEnds);
            builder.AppendFormat("Beam Width [cm] : {0}\n", BeamWidth);
            builder.AppendFormat("Beaming Tension [N/cm] : {0}\n", BeamingTension);
            builder.AppendFormat("Warping Tension [N/cm]: {0}\n", WarpingTension);
            builder.AppendFormat("Style No. : {0}\n", Style);

            return builder.ToString();
        }

        public void SetDefaultProperties()
        {
            DefaultStyle();
            DefaultFinalWidth();
            DefaultNumberOfEnds();
            DefaultSWrap();
            DefaultBeamWidth();
            DefaultConstants();
            DefaultTension();
            DefaultSpeedAndCrossWinding();         
            
        }

        private void DefaultStyle()
        {
            Style = StyleSpec.Style;
            Dtex = StyleSpec.Dtex;
        }

        private void DefaultFinalWidth()
        {
            if (Utils.Right(MaterialDescription, 2) == "CM")
            {
                FinalWidthCm = (float)Math.Round(Convert.ToDouble(Utils.Left(Utils.Right(MaterialDescription, 5), 3)), 0);
            }
            else
            {
                if (Utils.Left(MaterialDescription, 1) == null)
                {
                    FinalWidthCm = (float)Math.Round(Convert.ToDouble(Utils.Right(Utils.Left(MaterialDescription, 3), 2)), 0);
                }
                else
                {
                    throw new NotImplementedException();
                }
            }
        }

        private void DefaultNumberOfEnds()
        {
            float number_of_ends = (float)Math.Round(FinalWidthCm * EndsPerInch / 2.54, 0);
            if (number_of_ends % 2 == 0)
            {
                NumberOfEnds = number_of_ends;
            }
            else
            {
                NumberOfEnds = number_of_ends + 1;
            }
        }

        private void DefaultSWrap()
        {
			IsSWrapped |= Dtex >= 3000;
        }

        private void DefaultBeamWidth()
        {
            if (Dtex <= 3000)
            {
                BeamWidth = ((FinalWidthCm * 10) - 3) / 10;
            }
            else
            {
                BeamWidth = ((FinalWidthCm * 10) - 8) / 10;
            }
        }

        private void DefaultConstants()
        {
            K1 = (float)((YarnSupplier == "Dupont") ? 0.25 : 0.15);
            K2 = K1 + 1;
        }

        private void DefaultTension()
        {
            WarpingTension = (float)Math.Round(Dtex * K1, 0);
            BeamingTension = (float)Math.Round(NumberOfEnds * WarpingTension * K2 / 100, 0);
        }

        private void DefaultSpeedAndCrossWinding()
        {
            WarpingSpeed = 300; // meters / min
            if (Dtex >= HighDtex)
            {
                BeamingSpeed = 80; // meters / min
                CrossWinding = 10;
            }
            else
            {
                BeamingSpeed = 120; // meters / min
                CrossWinding = 5;
            }
        }
    }
}
