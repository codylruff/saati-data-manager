/*
 * C#
 * User: CRuff
 * Date: 3/8/2019
 * Time: 3:48 PM
 * 
 * 
 */
using System;

namespace DM_Lib
{
	/// <summary>
	/// Description of SpecTemplate.
	/// </summary>
	public class SpecTemplate
	{
		public string JsonText { get; set; }
		public string SpecType { get; set; }
		public string Revision { get; set; }
		
		public SpecTemplate(string json_text, string spec_type, string revision)
		{
			this.JsonText = json_text;
			this.SpecType = spec_type;
			this.Revision = revision;
		}
	}
}
