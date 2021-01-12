using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace onSet
{
    public class DataEntry : IExcelWriteable
    {
        //Fields
        string Column;
        string Table;
        string Value;
        string PrimaryKey;

        //SQL
        static SqlConnection cnn = DBManagement.cnn;
        static SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommand command;
        string sql;

        //Constructor
        public DataEntry(string Table, string Value)
        {
            this.Table = Table;
            this.Value = Value;
        }

        public DataEntry(string Table, string Value, string PrimaryKey)
        {
            this.Table = Table;
            this.Value = Value;
            this.PrimaryKey = PrimaryKey;
        }

        //Methods
        public void Submit()
        {
            sql = string.Format("Insert into {0} ({1}) values({2})", Table, Column, Value);
            adapter.InsertCommand = new SqlCommand(sql, cnn);
            adapter.InsertCommand.ExecuteNonQuery();

            command.Dispose();
        }

        public override string ToString()
        {
            return Value;
        }
    }
}
