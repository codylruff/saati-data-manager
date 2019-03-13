/*
 * C#
 * User: CRuff
 * Date: 3/12/2019
 * Time: 9:10 AM
 * 
 * 
 */
using System;
using System.Collections.Generic;

namespace DM_Lib
{
	/// <summary>
	/// Description of IRecord.
	/// </summary>
	public interface IRecord
	{
		string JsonText { get; set; }
		string SpecType { get; set; }
		string TimeStamp { get; set; }
		string TimeStampString { get; set; }
		string MaterialId { get; set; }
		int Id { get; set; }
		string Revision { get; set; }
		List<string> ColumnNames { get; set; }
		List<dynamic> Data { get; set; }
	}
}
