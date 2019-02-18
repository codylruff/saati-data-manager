/*
 * C#
 * User: CRuff
 * Date: 2/14/2019
 * Time: 3:39 PM
 * 
 * 
 */
using System.Runtime.InteropServices;
using ExcelDna.ComInterop;
using ExcelDna.Integration;
using DM_Lib;

namespace DM_Lib
{
	[ComVisible(true)]
	[ClassInterface(ClassInterfaceType.AutoDual)]
	public class DmComServer
	{
		public string GetSpecJson(string material_id)
		{
			var manager = new SpecManager();
			manager.LoadSpecification(material_id);
			string json_text = manager.SerializeSpec(manager.Specs.DefaultSpec);
			
			return json_text;
		}
		
		public long SendSpecJson(string json_text, string spec_type, string material_id)
		{
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
		public static object DnaComServerHello()
		{
			return "Hello from DnaComServer!";
		}
	}

}