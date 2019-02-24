/*
 * C#
 * User: CRuff
 * Date: 2/14/2019
 * Time: 3:39 PM
 * DM_Lib.DmComServer
 * 
 */
using System.Runtime.InteropServices;
using System.Collections.Generic;
using Newtonsoft.Json;
using ExcelDna.ComInterop;
using ExcelDna.Integration;
using DM_Lib;

namespace DM_Lib
{
	[ComVisible(true)]
	[ClassInterface(ClassInterfaceType.AutoDual)]
	public class DmComServer
	{
		public string GetStandardJson(string material_id)
		{
			var manager = new SpecManager();
			manager.LoadStandard(material_id);
			return manager.SerializeSpec(manager.Specs.DefaultSpec);
		}
		
		public string GetSpecJson(string material_id)
		{
			Dictionary<string,string> json_dict = new Dictionary<string,string>();
			
			var manager = new SpecManager();
			manager.LoadSpecification(material_id);
			foreach(ISpec spec in manager.Specs)
			{
				json_dict.Add(spec.Revision, manager.SerializeSpec(spec));
			}
			
			return JsonConvert.SerializeObject(json_dict);
		}
		
		public long PushSpecJson(string json_text, 
		                         string spec_type, string material_id, string revision, bool is_standard)
		{
			try{
				var record = new SpecRecord(json_text, spec_type, material_id, revision);
				var manager = new SpecManager();
				manager.CommitSpecificationRecord(
					record, is_standard ? "standard_specifications" : "modified_specifications");
			}
			catch{
				return -1;
			}
			return 0;
		}
	}

	[ComVisible(false)]
	public class ExcelAddin : IExcelAddIn
	{
		public void AutoOpen()
		{
			ComServer.DllRegisterServer();
		}
		public void AutoClose()
		{
			ComServer.DllUnregisterServer();
		}
	}
	
	public static class Functions	
	{
		[ExcelFunction]
		public static object DmComServerHello()
		{
			return "Hello from DmComServer!";
		}
	}

}