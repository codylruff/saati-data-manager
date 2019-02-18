/*
 * C#
 * User: CRuff
 * Date: 2/14/2019
 * Time: 11:08 AM
 * 
 * 
 */
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace DM_Lib
{
    public static class Factory
    {
        public static SpecRecord CreateSpecRecord()
        {
            var record = new SpecRecord();
            return record;
        }

        public static SpecRecord CreateRecordFromSpec(ISpec spec)
        {   
            string json_text = JsonConvert.SerializeObject(spec);
            var record = new SpecRecord(spec, json_text);
            return record;
        }

        public static SpecRecord CreateSpecRecordFromReader(SQLiteDataReader reader)
        {
            var record = new SpecRecord(reader);
            return record;
        }

        public static SpecRecord CreateSpecRecordFromList(List<string> list)
        {
            var record = new SpecRecord(list);
            return record;
        }

        public static ISpec CreateSpecFromRecord(SpecRecord record)
        {
            return DeserializeSpecification(record);
        }
        public static WarpingSpecification DefaultWarpingSpecification(string material_id, StyleSpecification style_spec)
        {
            var spec = new WarpingSpecification(material_id, style_spec);
            return spec;
        }

        public static WarpingSpecification WarpingSpecificationFromJson(string json_text)
        {
            var spec = new WarpingSpecification(json_text);
            return spec;
        }
        
        private static ISpec DeserializeSpecification(SpecRecord record)
        {
            switch (record.SpecType)
            {
                case "warping":
                    return JsonConvert.DeserializeObject<WarpingSpecification>(record.JsonText);
                case "style":
                    return JsonConvert.DeserializeObject<StyleSpecification>(record.JsonText);
                default:
                    return null;
            }
        }

    }
}