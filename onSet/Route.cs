using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace onSet
{
    public class Route : IExcelWriteable
    {
        //Fields
        public string color { get; private set; }
        public string wall { get; private set; }
        public string setter { get; private set; }

        public int SetterID { get; private set; }

        public string grade { get; private set; }
        public int Row { get; set; }

        public string PrimaryKey { get; private set; }

        //SQL
        static SqlConnection cnn = DBManagement.cnn;
        static SqlDataAdapter adapter = DBManagement.adapter;
        SqlCommand command;
        static SqlDataReader dataReader = DBManagement.dataReader;
        string sql;


        //Constructors
        public Route(string grade, string wall, string setter, string color)
        {
            this.grade = grade;
            this.color = color;
            this.wall = wall;
            this.setter = setter;
            sql = string.Format("SELECT SetterId" +
                "FROM Setters" +
                "WHERE SetterName = '{0}'", setter);
            command = new SqlCommand(sql, cnn);
            dataReader = command.ExecuteReader();

            while (dataReader.Read())
            {
                this.SetterID = (int) dataReader.GetValue(0);
            }
        }

        public Route(string grade, string wall, string setter, string color, string PrimaryKey)
        {
            this.grade = grade;
            this.color = color;
            this.wall = wall;
            this.setter = setter;
            this.PrimaryKey = PrimaryKey;
        }

        //Methods
        public virtual void Submit()
        {
            sql = string.Format("Insert into Routes (Wall, RouteGrade, SetterId, Color) values('{0}', '{1}', {2}, '{3}', '{4}')", wall, grade, SetterID.ToString(), color);
            command = new SqlCommand(sql, cnn);
            adapter.InsertCommand = command;
            adapter.InsertCommand.ExecuteNonQuery();

            command.Dispose();
        }

        public string ToString()
        {
            return string.Format("{0}, {1} {2}, {3}", setter, color, grade, wall);
        }

        public void SetRow(object sender, int Row)
        {
            if(sender is DBManagement)
            {
                this.Row = Row;
            }
        }
    }
}
