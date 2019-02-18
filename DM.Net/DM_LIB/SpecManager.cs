/*
 * C#
 * User: CRuff
 * Date: 2/14/2019
 * Time: 11:08 AM
 * 
 * 
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace DM_Lib
{
    public class SpecManager
    {
        public SpecsCollection Specs { get; set; }
        public List<string> MaterialsList { get; set; }

        public SpecManager()
        {
            Specs = new SpecsCollection();
            PopulateMaterialsList();
        }

        public void CreateNewMaterial(string material_id, string spec_type)
        {
            throw new NotImplementedException();
        }

        public void PrintSpecification()
        {
            Console.WriteLine("Standard Specification : \n{0}", Specs.DefaultSpec.ToString());
            foreach(ISpec spec in Specs)
            {
                Console.WriteLine("Recent Specifications : {0} \n{1}", spec.TimeStamp.ToString(), spec.ToString());
                Console.WriteLine(" - Next Spec - ");
                Console.ReadLine();
            }
        }
        
        public int LoadSpecification(string material_id)
        {
            Specs.DefaultSpec = GetDefaultSpec(material_id);
            Console.WriteLine(Specs.DefaultSpec.ToString());
            if (Specs.DefaultSpec == null) return -1;
            List<SpecRecord> records = DataAccess.GetSpecRecords(material_id, "modified_specifications");

            foreach (var record in records)
            {
                Specs.Add(Factory.CreateSpecFromRecord(record));
            }

            return 0;
        }

        public void CommitSpecification(ISpec spec)
        {
            DataAccess.PushSpec(spec);
        }
        
        public string SerializeSpec(ISpec spec)
        {
        	return JsonConvert.SerializeObject(spec);
        }

        private string GetPrintableSpecification(ISpec spec)
        {   
            return spec.ToString();
        }

        private ISpec GetDefaultSpec(string material_id)
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
        
        
    }
}