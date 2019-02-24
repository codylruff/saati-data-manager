/*
 * C#
 * User: CRuff
 * Date: 2/14/2019
 * Time: 11:07 AM
 * DM_Lib.DataAccess
 * 
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data;
using Newtonsoft.Json;

namespace DM_Lib
{
    public static class DataAccess
    {
        private const string connection_string = 
            @"Data Source=C:\Users\cruff\source\Spec Manager - COM\Database\SAATI_Spec_Manager.db3;Version=3;";

        public static SpecRecord SelectSingleRecord(string table_name, string field_name, dynamic field_value)
        {
            SQLiteDataReader reader = ExecuteSqlSelect(SqlSelectBuilder(table_name, field_name, field_value, 1));
            return Factory.CreateSpecRecordFromReader(reader);
        }

        public static List<SpecRecord> GetSpecRecords(string material_id, string table_name)
        {
            // Create a list of the records from the table
            SQLiteDataReader reader = ExecuteSqlSelect(SqlSelectBuilder(table_name, "Material_Id", material_id));
            List<SpecRecord> records = Factory.CreateListFromReader(reader);
            return records;
        }

        public static void PushSpec(string table_name, SpecRecord record = null, ISpec spec = null)
        {
            if(record == null && spec != null){
                record = Factory.CreateRecordFromSpec(spec);
                ExecuteSqlInsert(SqlInsertBuilder(table_name, record));
            }
            else if (spec == null && record != null){
                ExecuteSqlInsert(SqlInsertBuilder(table_name, record));
            }
            else{
                throw new System.ArgumentException("Must pass either a SpecRecord OR ISpec object");
            }
        }

        public static SQLiteDataReader ExecuteSqlSelect(string sql)
        {
            var dbConnection = new SQLiteConnection(connection_string);
            dbConnection.Open();        
            SQLiteCommand command = new SQLiteCommand(sql, dbConnection);
            return command.ExecuteReader();
        }

        public static void ExecuteSqlInsert(string sql)
        {
            var dbConnection = new SQLiteConnection(connection_string);
            dbConnection.Open();   
            SQLiteCommand command = new SQLiteCommand(sql, dbConnection);
            command.ExecuteNonQuery();
            dbConnection.Close();
        }


        private static string SqlSelectBuilder(string table_name, 
                                               string field_name, dynamic field_value, int limit = 0)
        {
            var sql = new StringBuilder();
            sql.AppendFormat("SELECT * FROM {0} WHERE {1} = '{2}'", table_name, field_name, field_value);
            return sql.ToString();
        }

        private static string SqlInsertBuilder(string table_name, SpecRecord record)
        {
            var sql = new StringBuilder();
            string insert = "INSERT INTO " + table_name + "(Material_Id, Time_Stamp, Spec_Type, Json_Text, Revision)";
            sql.AppendFormat("{0} VALUES ('{1}', '{2}', '{3}', '{4}', '{5}')",
                             insert, record.MaterialId, record.TimeStampString, record.SpecType, record.JsonText, record.Revision);
            return sql.ToString();
        }
    }
}
