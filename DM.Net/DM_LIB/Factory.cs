/*
 * C#
 * User: CRuff
 * Date: 2/14/2019
 * Time: 11:08 AM
 * DM_Lib.Factory
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
        	ISpec spec;
            switch (record.SpecType) 
			{
				case "warping":
            		spec = JsonConvert.DeserializeObject<WarpingSpecification>(record.JsonText);
            		spec.Revision = record.Revision;
            		spec.MaterialId = record.MaterialId;
            		spec.TimeStamp = record.TimeStamp;
            		return spec;
				case "style":
					spec = JsonConvert.DeserializeObject<StyleSpecification>(record.JsonText);
					spec.Revision = record.Revision;
            		spec.MaterialId = record.MaterialId;
            		spec.TimeStamp = record.TimeStamp;
            		return spec;
				default:
					throw new NotImplementedException();
    		}
        }
        

        public static ISpec CreateNewSpec(string material_id, string spec_type)
        {
            switch(spec_type)
            {
                case "warping":
                    return CreateDefaultWarpingSpecification(
            			material_id, CreateStyleFromNumber(Utils.Mid(material_id, 5, 3)));
                case "style":
            		return CreateDefaultStyleSpecification(material_id);
                default:
                    throw new NotImplementedException();
            }
        }

        public static StyleSpecification CreateDefaultStyleSpecification(string material_id)
        {
            throw new NotImplementedException();
        }

        public static ISpec CreateStyleFromNumber(string material_id)
        {
        	return CreateSpecFromRecord(DataAccess.SelectSingleRecord("standard_specifications", 
        	                                                          "Material_Id", material_id));
        }

        public static WarpingSpecification CreateDefaultWarpingSpecification(
                string material_id, ISpec style_spec)
        {
        	StyleSpecification style = (StyleSpecification)style_spec;
            var spec = new WarpingSpecification(material_id, style);
            return spec;
        }

        public static WarpingSpecification CreateWarpingSpecificationFromJson(string json_text)
        {
            var spec = new WarpingSpecification(json_text);
            return spec;
        }

        public static List<SpecRecord> CreateListFromReader(SQLiteDataReader reader)
        {
            List<SpecRecord> records = new List<SpecRecord>();
            Console.WriteLine(records.Count);
            List<string> fields = new List<string>();
            while (reader.Read())
            {	
                fields.Add((string)reader["Json_Text"]);
                fields.Add((string)reader["Spec_Type"]);
                fields.Add((string)reader["Material_Id"]);
                fields.Add(reader.GetInt32(0).ToString());
                fields.Add((string)reader["Revision"]);
                fields.Add((string)reader["Time_Stamp"]);
                records.Add(Factory.CreateSpecRecordFromList(fields));
                Console.WriteLine(records.Count);
            }
            
            return records;
        }
        
        public static void CreateSqliteDatabase(string db_name)
        {
            SQLiteConnection.CreateFile(db_name + ".sqlite");
        }
    }
}