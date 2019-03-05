/*
 * C#
 * User: CRuff
 * Date: 2/14/2019
 * Time: 11:08 AM
 * DM_Lib.SpecManager
 * 
 */
using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace DM_Lib
{
    public class SpecManager
    {
        public SpecsCollection Specs { get; set; }
        public SpecsCollection Standards { get; set; }
        public List<string> MaterialsList { get; set; }
        public string CorrectId { get; set; }

        public SpecManager()
        {
            Specs = new SpecsCollection();
            Standards = new SpecsCollection();
            PopulateMaterialsList();
        }

        public void CreateNewMaterial(string material_id, string spec_type)
        {
            var spec = Factory.CreateNewSpec(material_id, spec_type);
            
        }

        public void PrintSpecification()
        {
            Console.WriteLine("Standard Specification : \n{0}", Specs.DefaultSpec.ToString());
            Console.WriteLine("Recent Specifications : \n");
            foreach(ISpec spec in Specs)
            {
            	if(spec != Specs.DefaultSpec){
                Console.WriteLine("{0} \n{1}", spec.TimeStamp.ToString(), spec.ToString());
                Console.WriteLine(" - Next Spec - ");
                Console.ReadLine();
            	}
            }
        }
        
        public int LoadStandard(string material_id)
        {
        	
        	if((material_id != "101") && (Utils.Mid(material_id, 5, 3) != "101"))
        	{
        		Console.WriteLine("Debug");
        		Specs.DefaultSpec = GetDefaultSpec(material_id);
        	}else{
        		Console.WriteLine(material_id);
        		CorrectId = Handle101EdgeCase(material_id);
        		Specs.DefaultSpec = GetDefaultSpec(CorrectId);
        		return 1;
        	}
        	return Specs.DefaultSpec != null ? 0 : -1;
        }
        
        public void LoadSpecification(string material_id)
        {
            
            List<SpecRecord> records = DataAccess.GetSpecRecords(material_id, "modified_specifications");
            foreach (var record in records)
            {
                Specs.Add(Factory.CreateSpecFromRecord(record));
            }

        }

        public void CommitSpecificationRecord(SpecRecord record, string table_name)
        {
            DataAccess.PushSpec(table_name, record);
        }
        
        public string SerializeSpec(ISpec spec)
        {
        	return JsonConvert.SerializeObject(spec);
        }

        private string GetPrintableSpecification(ISpec spec)
        {   
            return spec.ToString();
        }

        public ISpec GetDefaultSpec(string material_id)
        {
            SpecRecord record = DataAccess.SelectSingleRecord("standard_specifications", "Material_Id", material_id);
            return Factory.CreateSpecFromRecord(record);
        }

        private void PopulateMaterialsList()
        {
            MaterialsList = new List<string>();
            MaterialsList.Add("warping");
            MaterialsList.Add("style");
        }
        
        private string Handle101EdgeCase(string material_id)
        {
        	if(material_id.Length >= 5){
        		return Utils.Mid(material_id, 5, 3) + Utils.Mid(material_id, 2, 2);
        	}else{
        		Console.WriteLine(@"Type 'KE' for Dupont yarn or 'HY' for Hyosung");
        		string input = Console.ReadLine();
        		return "101" + input;
        	}
        }
        
        
    }
}