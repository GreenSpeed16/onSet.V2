using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Charting = System.Windows.Forms.DataVisualization.Charting;
using Gui = System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace onSet
{
    public class Controller
    {
        //Fields
        DBManagement data;
        Form1 view;

        string Query;

        //Constructor
        public Controller(DBManagement d, Form1 v)
        {
            data = d;
            view = v;
        }

        //Methods
        public void UpdateGraph(string routeType, Charting.Chart chart)
        {
            string Query = "";
            string DataQuery = "";
            List<string> gradeLabels;

            if(routeType == "Boulder")
            {
                Query = "SELECT SUM(CASE WHEN RouteGrade = 'V0' THEN 1 ELSE 0 END) V0," +
                "SUM(CASE WHEN RouteGrade = 'V1' THEN 1 ELSE 0 END) V1," +
                "SUM(CASE WHEN RouteGrade = 'V2' THEN 1 ELSE 0 END) V2," +
                "SUM(CASE WHEN RouteGrade = 'V3' THEN 1 ELSE 0 END) V3," +
                "SUM(CASE WHEN RouteGrade = 'V4' THEN 1 ELSE 0 END) V4," +
                "SUM(CASE WHEN RouteGrade = 'V5' THEN 1 ELSE 0 END) V5," +
                "SUM(CASE WHEN RouteGrade = 'V6' THEN 1 ELSE 0 END) V6," +
                "SUM(CASE WHEN RouteGrade = 'V7' THEN 1 ELSE 0 END) V7," +
                "SUM(CASE WHEN RouteGrade = 'V8' THEN 1 ELSE 0 END) V8," +
                "SUM(CASE WHEN RouteGrade = 'V9' THEN 1 ELSE 0 END) V9" +
                "FROM Routes";

                DataQuery = "SELECT Goal" +
                    "FROM Goal" +
                    "WHERE Id <= 9";

                gradeLabels = DBManagement.boulderGrades;
            }
            else
            {
                Query = "SELECT SUM(CASE WHEN RouteGrade = '5.6' THEN 1 ELSE 0 END) 5.6," +
                "SUM(CASE WHEN RouteGrade = '5.11+' THEN 1 ELSE 0 END) 5.11+," +
                "SUM(CASE WHEN RouteGrade = '5.8' THEN 1 ELSE 0 END) 5.8," +
                "SUM(CASE WHEN RouteGrade = '5.9' THEN 1 ELSE 0 END) 5.9," +
                "SUM(CASE WHEN RouteGrade = '5.10-' THEN 1 ELSE 0 END) 5.10-," +
                "SUM(CASE WHEN RouteGrade = '5.10+' THEN 1 ELSE 0 END) 5.10+," +
                "SUM(CASE WHEN RouteGrade = '5.11-' THEN 1 ELSE 0 END) 5.11-," +
                "SUM(CASE WHEN RouteGrade = '5.11+' THEN 1 ELSE 0 END) 5.11+," +
                "SUM(CASE WHEN RouteGrade = '5.12-' THEN 1 ELSE 0 END) 5.12-," +
                "SUM(CASE WHEN RouteGrade = '5.12+' THEN 1 ELSE 0 END) 5.12+" +
                "SUM(CASE WHEN RouteGrade = '5.13-' THEN 1 ELSE 0 END) 5.13-" +
                "SUM(CASE WHEN RouteGrade = '5.13+' THEN 1 ELSE 0 END) 5.13+" +
                "FROM Routes";

                DataQuery = "SELECT Goal" +
                    "FROM Goal" +
                    "WHERE Id > 9";

                gradeLabels = DBManagement.ropeGrades;
            }

            data.UpdateGraph(Query, DataQuery, chart, gradeLabels);
        }

        public void SubmitParams(string paramText, string table, List<string> optionsList)
        {
            if(optionsList == DBManagement.wallOptions)
            {
                Query = "INSERT INTO {0}(WallName, WallType)" +
                    "VALUES ('{1}', 'Rope')";
            }
            else if(optionsList == DBManagement.ropeWallOptions)
            {
                Query = "INSERT INTO {0}(WallName, WallType)" +
                    "VALUES ('{1}', 'Boulder')";
            }
            else
            {
                Query = "INSERT INTO {0}" +
                    "VALUES ('{1}')";
            }

            data.SubmitParams(paramText, table, optionsList, Query);
        }

        public void UpdateOptions()
        {
            data.UpdateOptions();
        }

        public void ListView(string findValue, string routeFirstChar, Gui.CheckedListBox listView, bool isData, bool isWall, string wallType)
        {
            string Query = "";
            List<IExcelWriteable> CurrentList;

            if (isData)
            {
                if (isWall)
                {
                    Query = "SELECT WallName" +
                        "FROM Walls" +
                        string.Format("WHERE WallType = '{0}'", wallType);
                }
                else
                {
                    Query = string.Format("SELECT *" +
                        "FROM {0}", findValue);
                }

                data.ListData(Query, findValue, listView);
            }
            else
            {
                Query = "SELECT RouteGrade, Wall, Color, SetterName, RouteId" +
                "FROM Routes AS r JOIN Setters AS s" +
                "ON r.SetterId = s.SetterId" +
                string.Format("WHERE SUBSTRING(r.RouteGrade, 1, 1) = '{0}' AND r.Wall = '{1}'", routeFirstChar, findValue);

                //Find which route list to use
                if (routeFirstChar == "5")
                {
                    CurrentList = DBManagement.RopeList;
                }
                else
                {
                    CurrentList = DBManagement.BoulderList;
                }

                data.ListView(findValue, Query, listView, CurrentList);
            }
        }

        public void ReceiveGoal(string dataType, Gui.TabPage page)
        {
            List<int> primKeys;
            Query = "UPDATE Goal" +
                    "Goal = {0}" +
                    "WHERE Id = {1}";

            if(dataType == "Boulder")
            {
                primKeys = new List<int>()
                {
                    0,
                    1,
                    2,
                    3,
                    4,
                    5,
                    6,
                    7,
                    8,
                    9
                };
            }
            else
            {
                primKeys = new List<int>()
                {
                    10,
                    11,
                    12,
                    13,
                    14,
                    15,
                    16,
                    17,
                    18,
                    19,
                    20,
                    21
                };
            }

            data.ReceiveGoal(Query, page, primKeys);
        }

        public void Delete(Gui.CheckedListBox listView, string Type, string Table, string KeyColumn)
        {
            if(Type == "Rope")
            {
                data.DeleteRoutes(listView, DBManagement.RopeList, Table, KeyColumn);
            }
            else if(Type == "Boulder")
            {
                data.DeleteRoutes(listView, DBManagement.BoulderList, Table, KeyColumn);
            }
            else if(Type == "Data")
            {
                if(Table == "")
                data.DeleteRoutes(listView, DBManagement.DataList, Table, KeyColumn);
            }
        }

        public void SubmitRoute(string grade, string wall, string setter, string color)
        {
            data.SubmitRoute(new Route(grade, wall, setter, color));
        }

        public void Close()
        {
            data.Close();
        }
    }
}
