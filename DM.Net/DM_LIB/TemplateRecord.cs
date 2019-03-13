/*
 * C#
 * User: CRuff
 * Date: 3/12/2019
 * Time: 9:17 AM
 * 
 * 
 */
using System;
using System.Collections.Generic;

namespace DM_Lib
{
	/// <summary>
	/// Description of TemplateRecord.
	/// </summary>
	public class TemplateRecord : IRecord
	{
		public string JsonText { get; set; }
        public string SpecType { get; set; }
        public string TimeStamp { get; set; }
        public string TimeStampString { get; set; }
        public string MaterialId { get; set; }
        public int Id { get; set; }
        public string Revision { get; set; }
        public List<string> ColumnNames { get; set; }
        public List<dynamic> Data { get; set; }

        public TemplateRecord()
        {
            // Default constructor
        }
        
        public TemplateRecord(List<string> fields)
        {
        	JsonText = fields[0];
            SpecType = fields[1];
            Id = Convert.ToInt32(fields[2]);
            Revision = fields[3];
            TimeStampString = fields[4];           
        }
        
        public TemplateRecord(SpecTemplate template, string json_text)
        {
            JsonText = json_text;
            SpecType = template.SpecType;
            TimeStamp = DateTime.Now.ToString();
            Revision = template.Revision;
        }
	}
}
