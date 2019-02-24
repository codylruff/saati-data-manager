/*
 * C#
 * User: CRuff
 * Date: 2/14/2019
 * Time: 11:10 AM
 * DM_Lib.SpecRecord
 * 
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Globalization;

namespace DM_Lib
{
    public class SpecRecord
    {
        public string JsonText { get; private set; }
        public string SpecType { get; private set; }
        public DateTime TimeStamp { get; private set; }
        public string TimeStampString { get; private set; }
        public string MaterialId { get; private set; }
        public int Id { get; private set; }
        public string Revision { get; set; }

        public SpecRecord()
        {
            // Default constructor
        }

        public SpecRecord(List<string> fields)
        {
            JsonText = fields[0];
            SpecType = fields[1];
            MaterialId = fields[2];
            Id = Convert.ToInt32(fields[3]);
            Revision = fields[4];
            TimeStampString = fields[5];           
        }

        public SpecRecord(SQLiteDataReader reader)
        {
            while (reader.Read())
            {
                JsonText = (string)reader["Json_Text"];
                SpecType = (string)reader["Spec_Type"];
                MaterialId = (string)reader["Material_Id"];
                Revision = (string)reader["Revision"];
                Id = reader.GetInt32(0);
            }
        }

        public SpecRecord(ISpec spec, string json_text)
        {
            JsonText = json_text;
            SpecType = spec.SpecType;
            TimeStamp = DateTime.Now;
            MaterialId = spec.MaterialId;
            Revision = spec.Revision;
        }
        
        public SpecRecord(string json_text, string spec_type, string material_id, string revision)
        {
        	JsonText = json_text;
        	SpecType = spec_type;
            TimeStamp = DateTime.Now;
            TimeStampString = TimeStamp.ToString();
            MaterialId = material_id;
            Revision = revision;
        }
    }
}