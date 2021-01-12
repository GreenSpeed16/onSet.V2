using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Charting = System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;
using Gui = System.Windows.Forms;

namespace onSet
{
    public partial class Form1 : Form
    {
        //Initialize components
        Controller controller;
        private bool firstTime = true;

        //TextInfo for capitalization
        TextInfo myTI = new CultureInfo("en-US", false).TextInfo;

        //Route or boulder enum
        enum routeOrBoulder
        {
            Route,
            Boulder,
        }

        //Create route and preset list
        List<Route> routes = new List<Route>();

        //Constructor
        public Form1()
        {
            InitializeComponent();

            //Setup graph data
            ropeChart.ChartAreas[0].AxisX.Interval = 1;
            ropeChart.ChartAreas[0].AxisY.Interval = 5;
            ropeChart.Series[0].IsValueShownAsLabel = true;
            ropeChart.Series[1].IsValueShownAsLabel = true;

            boulderChart.ChartAreas[0].AxisX.Interval = 1;
            boulderChart.ChartAreas[0].AxisY.Interval = 5;
            boulderChart.Series[0].IsValueShownAsLabel = true;
            boulderChart.Series[1].IsValueShownAsLabel = true;

            if (firstTime) tabControl1.SelectTab(2);
            else tabControl1.SelectTab(0);
        }

        public void SetController(Controller e)
        {
            controller = e;
        }

        public void Setup()
        {
            controller.UpdateGraph("Rope", ropeChart);
            controller.UpdateGraph("Boulder", boulderChart);
            controller.UpdateOptions();
            ResetBoxes(boulderGradeBox, boulderColorBox, boulderWallBox, boulderSetterBox);
            ResetBoxes(ropeGradeBox, ropeColorBox, ropeWallBox, ropeSetterBox);
        }

        private void MassSubmit(Gui.RichTextBox gradeBox, Gui.RichTextBox colorBox, Gui.RichTextBox wallBox, Gui.RichTextBox setterBox, Worksheet routeSheet)
        {
            //Find last real row
            int lastUsedRow = routeSheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            int count = lastUsedRow + 1;

            if (gradeBox.Lines.Count() == colorBox.Lines.Count() && gradeBox.Lines.Count() == wallBox.Lines.Count() && gradeBox.Lines.Count() == setterBox.Lines.Count())
            {
                string sql = "";
                for (int i = 0; i < gradeBox.Lines.Count(); i++)
                {
                    routeSheet.Cells[count, "A"].Value = myTI.ToTitleCase(gradeBox.Lines[i]);
                    routeSheet.Cells[count, "B"].Value = myTI.ToTitleCase(wallBox.Lines[i]);
                    routeSheet.Cells[count, "C"].Value = myTI.ToTitleCase(setterBox.Lines[i]);
                    routeSheet.Cells[count, "D"].Value = myTI.ToTitleCase(colorBox.Lines[i]);

                    sql = "INSERT INTO Routes (Wall, RouteGrade, RouteSetter, SetterId, RouteType, Color) " +
                        string.Format("VALUES('{0}', '{1}', {2}, '{3}', '{4}')", wallBox.Lines[i], gradeBox.Lines[i], setterBox.Lines[i]);

                    count++;
                }

                MessageBox.Show("Routes submitted successfully");
                gradeBox.Clear();
                colorBox.Clear();
                wallBox.Clear();
                setterBox.Clear();
            }
            else
            {
                MessageBox.Show("Columns must have equal number of lines. (Check for accidental blank lines.)");
            }
        }

        //Event handlers
        public void ExcelManagement_NewGraphData(object sender, EventArgs e)
        {
            if(e is NewGraphDataEventArgs)
            {
                NewGraphDataEventArgs ev = e as NewGraphDataEventArgs;
                UpdateChart(ev.goalData, ev.currentData, ev.gradeLabels, ev.chart);
            }
        }

        internal void UpdateChart(List<int> goalList, List<int> routeList, List<string> labels, Charting.Chart chart)
        {
            //Set initial chart data
            chart.Series[0].Points.Clear();
            chart.Series[1].Points.Clear();

            //Assign new data
            for(int i = 0; i < labels.Count - 1; i++)
            {
                chart.Series[0].Points.AddXY(labels[i + 1], routeList[i]);
                chart.Series[1].Points.AddXY(labels[i + 1], goalList[i]);
            }

            //Update goal boxes
            if(goalList.Count == 10)
            {
                v0Box.Text = goalList[0].ToString();
                v1Box.Text = goalList[1].ToString();
                v2Box.Text = goalList[2].ToString();
                v3Box.Text = goalList[3].ToString();
                v4Box.Text = goalList[4].ToString();
                v5Box.Text = goalList[5].ToString();
                v6Box.Text = goalList[6].ToString();
                v7Box.Text = goalList[7].ToString();
                v8Box.Text = goalList[8].ToString();
                v9Box.Text = goalList[9].ToString();
            }
            else
            {
                ropeGoalBox01.Text = goalList[0].ToString();
                ropeGoalBox02.Text = goalList[1].ToString();
                ropeGoalBox03.Text = goalList[2].ToString();
                ropeGoalBox04.Text = goalList[3].ToString();
                ropeGoalBox05.Text = goalList[4].ToString();
                ropeGoalBox06.Text = goalList[5].ToString();
                ropeGoalBox07.Text = goalList[6].ToString();
                ropeGoalBox08.Text = goalList[7].ToString();
                ropeGoalBox09.Text = goalList[8].ToString();
                ropeGoalBox10.Text = goalList[9].ToString();
                ropeGoalBox11.Text = goalList[10].ToString();
                ropeGoalBox12.Text = goalList[11].ToString();
            }
        }

        public void ExcelManagement_PresetsUpdated(object sender, EventArgs e)
        {
            if(e is PresetsUpdatedEventArgs)
            {
                PresetsUpdatedEventArgs ev = e as PresetsUpdatedEventArgs;
                UpdateOptions(ev.setterOptions, ev.colorOptions, ev.wallOptions, ev.listWallOptions, ev.ropeWallOptions, ev.listRopeWallOptions, ev.listParamOptions
                    , ev.boulderGradeOptions, ev.ropeGradeOptions);
            }
        }

        internal void UpdateOptions(List<string> setterOptions, List<string> colorOptions, List<string> wallOptions, List<string> listWallOptions, 
            List<string> ropeWallOptions, List<string> listRopeWallOptions, List<string> listParamOptions, List<string> boulderGradeOptions, List<string> ropeGradeOptions)
        {
            //Update comboboxes
            boulderSetterBox.DataSource = setterOptions;
            ropeSetterBox.DataSource = setterOptions;

            boulderColorBox.DataSource = colorOptions;
            ropeColorBox.DataSource = colorOptions;

            boulderWallBox.DataSource = wallOptions;
            boulderDeleteSelect.DataSource = listWallOptions;
            ropeWallBox.DataSource = ropeWallOptions;
            ropeDeleteSelect.DataSource = listRopeWallOptions;

            presetBox.DataSource = listParamOptions;
            boulderGradeBox.DataSource = boulderGradeOptions;
            ropeGradeBox.DataSource = ropeGradeOptions;
        }

        public void ExcelManagement_NewListView(object sender, EventArgs e)
        {
            if (e is NewListViewEventArgs)
            {
                NewListViewEventArgs ev = e as NewListViewEventArgs;
                ListView(ev.DataList, ev.ListView);
            }
        }

        public void ListView(List<IExcelWriteable> List, CheckedListBox ListView)
        {
            ListView.Items.Clear();

            foreach (IExcelWriteable Item in List)
            {
                ListView.Items.Add(Item);
            }
        }

        //Old Below this point

        private void ResetBoxes(Gui.ComboBox gradeBox, Gui.ComboBox colorBox, Gui.ComboBox wallBox, Gui.ComboBox setterBox)
        {
            gradeBox.SelectedIndex = gradeBox.FindStringExact("Grade:");
            colorBox.SelectedIndex = colorBox.FindStringExact("Color:");
            wallBox.SelectedIndex = wallBox.FindStringExact("Wall:");
            setterBox.SelectedIndex = setterBox.FindStringExact("Setter:");
        }       

        //Guarantee that all options are selected before enabling submission
        private void boulderSetterBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (boulderGradeBox.Text != "Grade:" && boulderColorBox.Text != "Color:" && boulderWallBox.Text != "Wall:" && boulderSetterBox.Text != "Setter:") 
                    boulderSubmitButton.Enabled = true;
            else boulderSubmitButton.Enabled = false;
        }

        private void boulderColorBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (boulderGradeBox.Text != "Grade:" && boulderColorBox.Text != "Color:" && boulderWallBox.Text != "Wall:" && boulderSetterBox.Text != "Setter:")
                boulderSubmitButton.Enabled = true;
            else boulderSubmitButton.Enabled = false;
        }

        private void boulderWallBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (boulderGradeBox.Text != "Grade:" && boulderColorBox.Text != "Color:" && boulderWallBox.Text != "Wall:" && boulderSetterBox.Text != "Setter:")
                boulderSubmitButton.Enabled = true;
            else boulderSubmitButton.Enabled = false;
        }

        private void boulderGradeBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (boulderGradeBox.Text != "Grade:" && boulderColorBox.Text != "Color:" && boulderWallBox.Text != "Wall:" && boulderSetterBox.Text != "Setter:")
                boulderSubmitButton.Enabled = true;
            else boulderSubmitButton.Enabled = false;
        }

        private void ropeGradeBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ropeGradeBox.Text != "Grade:" && ropeColorBox.Text != "Color:" && ropeWallBox.Text != "Wall:" && ropeSetterBox.Text != "Setter:")
                ropeSubmitButton.Enabled = true;
            else ropeSubmitButton.Enabled = false;
        }

        private void ropeColorBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ropeGradeBox.Text != "Grade:" && ropeColorBox.Text != "Color:" && ropeWallBox.Text != "Wall:" && ropeSetterBox.Text != "Setter:")
                ropeSubmitButton.Enabled = true;
            else ropeSubmitButton.Enabled = false;
        }

        private void ropeWallBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ropeGradeBox.Text != "Grade:" && ropeColorBox.Text != "Color:" && ropeWallBox.Text != "Wall:" && ropeSetterBox.Text != "Setter:")
                ropeSubmitButton.Enabled = true;
            else ropeSubmitButton.Enabled = false;
        }

        private void ropeSetterBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ropeGradeBox.Text != "Grade:" && ropeColorBox.Text != "Color:" && ropeWallBox.Text != "Wall:" && ropeSetterBox.Text != "Setter:")
                ropeSubmitButton.Enabled = true;
            else ropeSubmitButton.Enabled = false;
        }

        //Submit Route
        private void boulderSubmitButton_Click(object sender, EventArgs e)
        {
            controller.SubmitRoute(boulderGradeBox.Text, boulderWallBox.Text, boulderSetterBox.Text, boulderColorBox.Text);
            ResetBoxes(boulderGradeBox, boulderColorBox, boulderWallBox, boulderSetterBox);
            controller.UpdateGraph("Boulder", boulderChart);
        }

        private void ropeSubmitButton_Click(object sender, EventArgs e)
        {
            controller.SubmitRoute(ropeGradeBox.Text, ropeWallBox.Text, ropeSetterBox.Text, ropeColorBox.Text);
            ResetBoxes(ropeGradeBox, ropeColorBox, ropeWallBox, ropeSetterBox);
            controller.UpdateGraph("Rope", ropeChart);
        }

        //Route list
        private void boulderDeleteSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            controller.ListView(boulderDeleteSelect.Text, "V", boulderDeleteBox, false, false, "Boulder");
        }

        private void boulderDelButton_Click(object sender, EventArgs e)
        {
            controller.Delete(boulderDeleteBox, "Boulder", "Routes", "RouteId");
            controller.UpdateGraph("Boulder", boulderChart);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            controller.ListView(ropeDeleteSelect.Text, "5", ropeDeleteList, false, false, "Rope");
        }

        private void ropeDeleteButton_Click(object sender, EventArgs e)
        {            
            controller.Delete(ropeDeleteList, "Rope", "Routes", "RouteId");
            controller.UpdateGraph("Rope", ropeChart);
        }

        private void presetBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool isWall = false;
            string Table;
            string KeyColumn;

            switch (presetBox.Text)
            {
                case "Setters":
                    Table = "Setters";
                    KeyColumn = "SetterId";
                    break;
                case "Colors":
                    Table = "Colors";
                    KeyColumn = "Color";
                    break;
                case "Boulder Walls":
                    isWall = true;
                    Table = "Walls";
                    KeyColumn = "WallName";
                    break;
                case "Rope Walls":
                    isWall = true;
                    break;
            }

            
            controller.ListView(presetBox.Text, "DATA", presetListBox, true, isWall, "None");
        }

        private void presetDeleteButton_Click(object sender, EventArgs e)
        {
            
            controller.Delete(presetListBox, "Data", presetBox.Text, "");
        }

        private void boulderCurveSubmitButton_Click(object sender, EventArgs e)
        {
            controller.ReceiveGoal("Boulder", curveControl.SelectedTab);
            controller.UpdateGraph("Boulder", boulderChart);
        }

        private void ropeCurveSubmit_Click(object sender, EventArgs e)
        {
            controller.ReceiveGoal("Rope", curveControl.SelectedTab);
            controller.UpdateGraph("Rope", ropeChart);
        }


        private void boulderMassSubmitButton_Click(object sender, EventArgs e)
        {
            MassSubmit(boulderGradeMassSubmit, boulderColorMassSubmit, boulderWallMassSubmit, boulderSetterMassSubmit);
            controller.UpdateGraph("Boulder", boulderChart);
        }

        private void ropeMassSubmitButton_Click(object sender, EventArgs e)
        {
            MassSubmit(ropeGradeMassSubmit, ropeColorMassSubmit, ropeWallMassSubmit, ropeSetterMassSubmit);
            controller.UpdateGraph("Rope", ropeChart);
        }

        private void parameterSubmitButton_Click(object sender, EventArgs e)
        {
            if(setterOptionsBox.Text != "") controller.SubmitParams(setterOptionsBox.Text, "Setters", DBManagement.setterOptions);

            if(colorOptionsBox.Text !="") controller.SubmitParams(colorOptionsBox.Text, "Colors", DBManagement.colorOptions);

            if(wallOptionsBox.Text != "") controller.SubmitParams(wallOptionsBox.Text, "Walls", DBManagement.wallOptions);

            if(rWallOptionsBox.Text != "") controller.SubmitParams(rWallOptionsBox.Text, "Walls", DBManagement.ropeWallOptions);

            controller.UpdateOptions();

            setterOptionsBox.Clear();
            colorOptionsBox.Clear();
            wallOptionsBox.Clear();
            rWallOptionsBox.Clear();

            MessageBox.Show("Options Submitted!");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            controller.Close();
        }
    }
}
