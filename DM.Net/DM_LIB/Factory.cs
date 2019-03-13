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

        public static SpecRecord CreateRecordFromTemplate(SpecTemplate template)
        {
            var record = new SpecRecord(template, template.JsonText);
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
            		spec.TimeStamp = Convert.ToDateTime(record.TimeStamp);
            		return spec;
				case "style":
					spec = JsonConvert.DeserializeObject<StyleSpecification>(record.JsonText);
					spec.Revision = record.Revision;
            		spec.MaterialId = record.MaterialId;
            		spec.TimeStamp = Convert.ToDateTime(record.TimeStamp);
            		return spec;
            	case "fabric":
            		spec = JsonConvert.DeserializeObject<FabricSpecification>(record.JsonText);
					spec.Revision = record.Revision;
            		spec.MaterialId = record.MaterialId;
            		spec.TimeStamp = Convert.ToDateTime(record.TimeStamp);
            		return spec;
            	case "generic":
            		spec = JsonConvert.DeserializeObject<GenericSpecification>(record.JsonText);
					spec.Revision = record.Revision;
            		spec.MaterialId = record.MaterialId;
            		spec.TimeStamp = Convert.ToDateTime(record.TimeStamp);
            		return spec;
				default:
					throw new NotImplementedException();
    		}
        }
        

        public static ISpec CreateNewSpec(string material_id, string spec_type, string json_text = null)
        {
            switch(spec_type)
            {
                case "warping":
                    return CreateDefaultWarpingSpecification(
            			material_id, CreateStyleFromNumber(Utils.Mid(material_id, 5, 3)));
                case "style":
            		return CreateDefaultStyleSpecification(material_id);
            	case "fabric":
            		return CreateDefaultFabricSpecification(
            			material_id, CreateStyleFromNumber(Utils.Mid(material_id, 5, 3)));
                default:
                    return CreateGenericSpecification(material_id, json_text);
            }
        }
        
        public static ISpec CreateGenericSpecification(string material_id, string json_text)
        {
        	var spec = new GenericSpecification(json_text);
        	spec.MaterialId = material_id;
        	return spec;
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
        
        public static FabricSpecification CreateDefaultFabricSpecification(string material_id, ISpec style_spec)
        {
        	StyleSpecification style = (StyleSpecification)style_spec;
        	var spec = JsonConvert.DeserializeObject<FabricSpecification>(JsonConvert.SerializeObject(style));
        	spec.MaterialId = material_id;
        	return spec;
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
            List<string> fields;
            while (reader.Read())
            {	
            	fields = Factory.CreateStringList();
                fields.Add((string)reader["Json_Text"]);
                fields.Add((string)reader["Spec_Type"]);
                fields.Add((string)reader["Material_Id"]);
                fields.Add(reader.GetInt32(0).ToString());
                fields.Add((string)reader["Revision"]);
                Console.WriteLine((string)reader["Revision"]);
                fields.Add((string)reader["Time_Stamp"]);
                records.Add(Factory.CreateSpecRecordFromList(fields));
                Console.WriteLine(records.Count);
                fields = null;
            }
            
            return records;
        }
        
        public static List<string> CreateStringList()
        {
        	return new List<string>();
        }
        
        public static void CreateSqliteDatabase(string db_name)
        {
            SQLiteConnection.CreateFile(db_name + ".sqlite");
        }
    }
}