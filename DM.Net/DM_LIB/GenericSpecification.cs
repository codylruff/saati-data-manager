/*
 * C#
 * User: CRuff
 * Date: 2/24/2019
 * Time: 4:31 PM
 * 
 * 
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace DM_Lib
{
	/// <summary>
	/// Description of GenericSpecification.
	/// </summary>
	public class GenericSpecification : ISpec
	{
		public string MaterialId { get; set; }
        public DateTime TimeStamp { get; set; }
        public string SpecType { get; set;}
        public bool IsDefault { get; set; }
        public string Revision { get; set;}
        public Dictionary<string,string> Properties { get; set; }
        
        [JsonIgnore]
        public ISpec ParentSpec { get; set; }
        
		public GenericSpecification()
		{
			// Default Constructor
		}
		
		public GenericSpecification(string json_text)
		{
			this.Properties = JsonConvert.DeserializeObject<Dictionary<string,string>>(json_text);
		}
		
        public void SetDefaultProperties()
        {
            // Currently these are not calculated but 
            // simply input into the database as data.
        }
        
        public override string ToString()
        {
            var builder = new StringBuilder();
            foreach(var kvp in Properties)
            {
            	builder.AppendFormat("{0} : {1}\n", kvp.Key, kvp.Value);
            }
            return builder.ToString();
        }
		
	}
}
