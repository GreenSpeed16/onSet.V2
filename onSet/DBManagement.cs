using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Charting = System.Windows.Forms.DataVisualization.Charting;
using Gui = System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Data.SqlClient;

namespace onSet
{
    public class DBManagement
    {
        //Constants
        public static List<string> boulderGrades { get; private set; }
        public static List<string> ropeGrades { get; private set; }

        //SQL
        static string connectionString;
        public static SqlConnection cnn { get; private set; }
        public static SqlDataAdapter adapter { get; private set; }

        SqlCommand command;
        public static SqlDataReader dataReader { get; private set; }
        string sql = "";

        //Option lists
        public static List<string> setterOptions { get; private set; }
        public static List<string> colorOptions { get; private set; }
        public static List<string> wallOptions { get; private set; }
        public static List<string> listWallOptions { get; private set; }
        public static List<string> ropeWallOptions { get; private set; }
        public static List<string> listRopeWallOptions { get; private set; }
        public static List<string> listParamOptions { get; private set; }

        //Event handlers
        public event EventHandler NewGraphData;
        public event EventHandler NewRouteAdded;
        public event EventHandler RouteDeleted;
        public event EventHandler PresetsUpdated;
        public event EventHandler GoalUpdated;
        public event EventHandler NewListView;

        //Program Information
        public static bool firstTime /*Open program to options tab on first use*/{ get; private set; }
        static int[] ropeRouteData = new int[10];
        static int[] boulderRouteData = new int[12];
        static int[] boulderGoalData = new int[10];
        static int[] ropeGoalData = new int[12];
        public static List<IExcelWriteable> BoulderList { get; private set; }
        public static List<IExcelWriteable> RopeList { get; private set; }
        public static List<IExcelWriteable> DataList { get; private set; }

        //TextInfo for capitalization
        static TextInfo myTI = new CultureInfo("en-US", false).TextInfo;

        //Constructor
        public DBManagement()
        {
            connectionString = @"Server=.;Database=gymDatabase;Trusted_Connection=True;";
            cnn = new SqlConnection(connectionString);
            cnn.Open();

            boulderGrades = new List<string>()
        {
            "Grade:",
            "V0",
            "V1",
            "V2",
            "V3",
            "V4",
            "V5",
            "V6",
            "V7",
            "V8",
            "V9"
        };
            ropeGrades = new List<string>()
        {
            "Grade:",
            "5.6",
            "5.7",
            "5.8",
            "5.9",
            "5.10-",
            "5.10+",
            "5.11-",
            "5.11+",
            "5.12-",
            "5.12+",
            "5.13-",
            "5.13+"
        };
            
            adapter = new SqlDataAdapter();
        }

        //Methods
        public void UpdateGraph(string Query, string DataQuery, Charting.Chart chart, List<string> gradeLabels)
        {
            List<int> gradeList = new List<int>();
            List<int> goalList = new List<int>();
            //Current Graph Data
            //Query
            command = new SqlCommand(Query, cnn);
            dataReader = command.ExecuteReader();

            //Read data
            while (dataReader.Read())
            {
                gradeList.Add((int)dataReader.GetValue(0));
            }

            command.Dispose();
            dataReader.Close();

            //Get data for goal graph
            //Query
            command = new SqlCommand(DataQuery, cnn);
            dataReader = command.ExecuteReader();

            //Read data
            while (dataReader.Read())
            {
                goalList.Add((int)dataReader.GetValue(0));
            }

            command.Dispose();
            dataReader.Close();

            EventHandler newGraphData = NewGraphData;
            if (newGraphData != null)
            {
                NewGraphDataEventArgs e = new NewGraphDataEventArgs();
                e.currentData = gradeList;
                e.goalData = goalList;
                e.chart = chart;
                e.gradeLabels = gradeLabels;

                NewGraphData(this, e);
            }
        }

        public void SubmitParams(string paramText, string table, List<string> optionsList, string Query)
        {
            //Params
            List<string> paramList;
            paramList = new List<string>(paramText.Split(','));

            foreach (string param in paramList)
            {
                sql = string.Format(Query, table, param);

                command = new SqlCommand(sql, cnn);

                try
                {
                    adapter.InsertCommand = command;
                    adapter.InsertCommand.ExecuteNonQuery();
                }
                catch(SqlException exception)
                {
                    if (exception.Number == 2601) continue; //Cannot insert duplicate
                    else throw; //Catch unexpected error
                }

                optionsList.Add(param);
            }
        }

        public void InitialOptions()
        {
            //Add initial options
            setterOptions.Add("Setter:");
            colorOptions.Add("Color:");
            wallOptions.Add("Wall:");
            listWallOptions.Add("Select:");
            ropeWallOptions.Add("Wall:");
            listRopeWallOptions.Add("Select:");

            //Get data
            //Setters
            sql = "SELECT SetterName" +
                "FROM Setters";
            command = new SqlCommand(sql, cnn);
            dataReader = command.ExecuteReader();

            while (dataReader.Read())
            {
                setterOptions.Add(dataReader.GetValue(0).ToString());
            }

            dataReader.Close();
            command.Dispose();

            //Colors
            sql = "SELECT Color" +
                "FROM Colors";
            command = new SqlCommand(sql, cnn);
            dataReader = command.ExecuteReader();

            while (dataReader.Read())
            {
                colorOptions.Add(dataReader.GetValue(0).ToString());
            }

            dataReader.Close();
            command.Dispose();

            //Boulder walls
            sql = "SELECT WallName" +
                "FROM Walls" +
                "WHERE WallType = 'Boulder'";
            command = new SqlCommand(sql, cnn);
            dataReader = command.ExecuteReader();

            while (dataReader.Read())
            {
                wallOptions.Add(dataReader.GetValue(0).ToString());
                listWallOptions.Add(dataReader.GetValue(0).ToString());
            }

            dataReader.Close();
            command.Dispose();

            //Rope walls
            sql = "SELECT WallName" +
                "FROM Walls" +
                "WHERE WallType = 'Rope'";
            command = new SqlCommand(sql, cnn);
            dataReader = command.ExecuteReader();

            while (dataReader.Read())
            {
                ropeWallOptions.Add(dataReader.GetValue(0).ToString());
                listRopeWallOptions.Add(dataReader.GetValue(0).ToString());
            }

            dataReader.Close();
            command.Dispose();

            //Event
            UpdateOptions();
        }

        public void UpdateOptions()
        {
            EventHandler presetsUpdated = PresetsUpdated;
            if (presetsUpdated != null)
            {
                PresetsUpdatedEventArgs e = new PresetsUpdatedEventArgs();
                e.setterOptions = setterOptions;
                e.colorOptions = colorOptions;
                e.wallOptions = wallOptions;
                e.listWallOptions = listWallOptions;
                e.ropeWallOptions = ropeWallOptions;
                e.listRopeWallOptions = listRopeWallOptions;
                e.boulderGradeOptions = boulderGrades;
                e.ropeGradeOptions = ropeGrades;
                e.listParamOptions = listParamOptions;

                presetsUpdated(this, e);
            }
        }

        public void ListView(string findValue, string Query, Gui.CheckedListBox listView, List<IExcelWriteable> CurrentList)
        {
            //Select query
            command = new SqlCommand(Query, cnn);
            dataReader = command.ExecuteReader();

            //Create routes
            while (dataReader.Read())
            {
                CurrentList.Add(new Route(dataReader.GetValue(0).ToString(), dataReader.GetValue(1).ToString(), dataReader.GetValue(2).ToString(), dataReader.GetValue(3).ToString(), dataReader.GetValue(4).ToString()));
            }

            //Dispose command and reader
            dataReader.Close();
            command.Dispose();

            //Convert list to IExcelWriteable
            List<IExcelWriteable> PassList = new List<IExcelWriteable>();

            foreach(Route route in CurrentList)
            {
                PassList.Add(route);
            }

            //Throw event
            NewListViewEventArgs e = new NewListViewEventArgs();
            e.DataList = PassList;
            e.ListView = listView;
            NewListView(this, e);
        }

        public void ListData(string Query, string Table, Gui.CheckedListBox listView)
        {
            command = new SqlCommand(Query, cnn);
            dataReader = command.ExecuteReader();

            //Create data
            while (dataReader.Read())
            {
                DataList.Add(new DataEntry(Table, dataReader.GetValue(1).ToString(), dataReader.GetValue(0).ToString()));
            }

            //Convert list to IExcelWriteable
            List<IExcelWriteable> PassList = new List<IExcelWriteable>();

            foreach (DataEntry Data in DataList)
            {
                PassList.Add(Data);
            }

            //Throw event
            NewListViewEventArgs e = new NewListViewEventArgs();
            e.DataList = PassList;
            e.ListView = listView;
            NewListView(this, e);
        }

        public void ReceiveGoal(string Query, Gui.TabPage page, List<int> primKeys)
        {
            //Fields
            List<Gui.TextBox> entryList = new List<Gui.TextBox>();

            //Query
            foreach (Gui.Control control in page.Controls)
            {
                if (control is Gui.TextBox)
                {
                    entryList.Add((Gui.TextBox)control);
                }
            }

            for(int i = 0; i < entryList.Count; i++)
            {
                if (int.TryParse(entryList[i].Text, out int dataInt))
                {
                    sql = string.Format(Query, primKeys[i], dataInt);
                    command = new SqlCommand(sql, cnn);
                    adapter.UpdateCommand = command;

                    adapter.UpdateCommand.ExecuteNonQuery();
                    command.Dispose();
                }
            }

            Gui.MessageBox.Show("Data Entered!");

            UpdateOptions();
        }

        public void DeleteRoutes(Gui.CheckedListBox listView, List<IExcelWriteable> CurrentList, string Table, string KeyColumn)
        {
            string SqlDelete = string.Empty;
            List<int> deleteList = new List<int>();
            
            //Create delete list
            for (int i = 0; i < CurrentList.Count; i++)
            {
                if (listView.GetItemChecked(i))
                {
                    deleteList.Add(int.Parse(CurrentList[i].PrimaryKey));
                }
            }

            //Delete routes
            for (int i = 0; i < deleteList.Count; i++)
            {
                SqlDelete += deleteList[i] + ",";
            }

            SqlDelete = SqlDelete.TrimEnd(',');

            command = new SqlCommand(string.Format("DELETE FROM {0}" +
                "WHERE {1} IN ({2})", Table, KeyColumn, SqlDelete), cnn);
            adapter.DeleteCommand = command;
            adapter.DeleteCommand.ExecuteNonQuery();

            command.Dispose();
        }

        public void SubmitRoute(Route route)
        {
            route.Submit();
        }

        public void Close()
        {
            cnn.Close();
        }
    }
}
