/*
 * C#
 * User: CRuff
 * Date: 2/14/2019
 * Time: 11:07 AM
 * 
 * 
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data;

namespace DM_Lib
{
    public static class DataAccess
    {
        public static SpecRecord SelectSingleRecord(string table_name, string field_name, dynamic field_value)
        {
            SQLiteDataReader reader = ExecuteSqlSelect(SqlSelectBuilder(table_name, field_name, field_value, 1));
            return Factory.CreateSpecRecordFromReader(reader);
        }

        public static List<SpecRecord> GetSpecRecords(string material_id, string table_name)
        {
            // Create a list of the records from the table
            SQLiteDataReader reader = ExecuteSqlSelect(SqlSelectBuilder(table_name, "Material_Id", material_id));
            List<SpecRecord> records = CreateListFromReader(reader);

            return records;
        }

        public static void PushSpec(ISpec spec)
        {
            var record = Factory.CreateRecordFromSpec(spec);
            // TODO: Create sql to load into table: modified_specificiations
        }

        public static SQLiteDataReader ExecuteSqlSelect(string sql)
        {
            var dbConnection = new SQLiteConnection(
                @"Data Source=C:\Users\cruff\source\Spec Manager - COM\Database\SAATI_Spec_Manager.db3;Version=3;");
            dbConnection.Open();        
            SQLiteCommand command = new SQLiteCommand(sql, dbConnection);
            return command.ExecuteReader();
        }

        public static void CreateSqliteDatabase(string db_name)
        {
            SQLiteConnection.CreateFile(db_name + ".sqlite");
        }

        private static string SqlSelectBuilder(string table_name, string field_name, dynamic field_value, int limit = 0)
        {
            var sql = new StringBuilder();
            sql.AppendFormat("SELECT * FROM {0} WHERE {1} = '{2}'", table_name, field_name, field_value);
            if (limit > 0) sql.AppendFormat("\n LIMIT {0}", limit);
            return sql.ToString();
        }

        private static List<SpecRecord> CreateListFromReader(SQLiteDataReader reader)
        {
            List<SpecRecord> records = new List<SpecRecord>();
            List<string> fields = new List<string>();
            while (reader.Read())
            {
                fields.Add((string)reader["Json_Text"]);
                Console.WriteLine(fields[0]);
                fields.Add((string)reader["Spec_Type"]);
                Console.WriteLine(fields[1]);
                fields.Add((string)reader["Material_Id"]);
                Console.WriteLine(fields[2]);
                fields.Add(reader.GetInt32(0).ToString());
                Console.WriteLine(fields[3]);
                records.Add(Factory.CreateSpecRecordFromList(fields));
            }
            
            return records;
        }

    }
}
