using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Chart = System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;
using Gui = System.Windows.Forms;

namespace onSet
{
    public partial class Form1 : Form
    {
        //Initialize components
        public Excel.Application reader = new Excel.Application();
        public Workbook routeBook;
        public Worksheet ropeSheet;
        public Worksheet boulderSheet;
        public Worksheet dataSheet;

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

        private void CheckForOpenFile()
        {
            
            //Open workbook
            try
            {
                Stream s = File.Open("Routes.xlsx", FileMode.Open, FileAccess.Read, FileShare.None);

                s.Close();

                routeBook = reader.Workbooks.Open(string.Format("{0}\\Routes.xlsx", System.IO.Directory.GetCurrentDirectory()));
                dataSheet = routeBook.Worksheets[3];
                ropeSheet = routeBook.Worksheets[2];
                boulderSheet = routeBook.Worksheets[1];

                firstTime = false;
            }
            catch (FileNotFoundException)
            {
                routeBook = reader.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                boulderSheet = routeBook.Worksheets[1];
                boulderSheet.Name = "Boulders";

                ropeSheet = routeBook.Worksheets.Add(After: boulderSheet);
                ropeSheet.Name = "Ropes";

                dataSheet = routeBook.Worksheets.Add(After: ropeSheet);
                dataSheet.Name = "Data";

                boulderSheet.Cells[1, "A"].Value = "Grade";
                boulderSheet.Cells[1, "B"].Value = "Wall";
                boulderSheet.Cells[1, "C"].Value = "Setter";
                boulderSheet.Cells[1, "D"].Value = "Color";

                ropeSheet.Cells[1, "A"].Value = "Grade";
                ropeSheet.Cells[1, "B"].Value = "Wall";
                ropeSheet.Cells[1, "C"].Value = "Setter";
                ropeSheet.Cells[1, "D"].Value = "Color";

                dataSheet.Cells[1, "A"].Value = "Boulder";
                dataSheet.Cells[1, "B"].Value = "Ropes";
                dataSheet.Cells[1, "C"].Value = "Setters";
                dataSheet.Cells[1, "D"].Value = "Colors";
                dataSheet.Cells[1, "E"].Value = "Walls";
                dataSheet.Cells[1, "F"].Value = "RWalls";

                routeBook.SaveAs(string.Format("{0}\\Routes.xlsx", System.IO.Directory.GetCurrentDirectory()));
            }
            catch (System.IO.IOException)
            {
                MessageBox.Show("Program cannot open if the spreadsheet is open elsewhere. Please close spreadsheet and try again.");
                Environment.Exit(0);
            }
            

            UpdateOptions();

            UpdateChart(boulderSheet, boulderChart, true);

            UpdateChart(ropeSheet, ropeChart, true);
        }

        private void ListRoutes(Gui.ComboBox selector, Gui.CheckedListBox listView, Worksheet routeSheet, string column, bool isPreset)
        {
            //Empty route lists
            listView.Items.Clear();

            if (isPreset) listView.Items.Add("Choose A Preset");
            else listView.Items.Add("No Routes");

            routes.Clear();

            //Fill route list
            //Find last real row
            int lastUsedRow = routeSheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            for (int i = 2; i <= lastUsedRow; i++)
            {
                if (isPreset)
                {
                    if (routeSheet.Cells[i, column].Value != null && routeSheet.Cells[i, column].Value.ToString() != "")
                    {
                        listView.Items.Remove("Choose A Preset");
                        Route newRoute = new Route(routeSheet.Cells[i, column].Value.ToString(), "Preset", "Preset", "Preset");
                        newRoute.Row = i;
                        newRoute.Column = column;
                        routes.Add(newRoute);
                        listView.Items.Add(newRoute);
                    }
                }
                else
                {
                    if ((string)routeSheet.Cells[i, "B"].Value == selector.Text)
                    {
                        listView.Items.Remove("No Routes");
                        Route newRoute = new Route(routeSheet.Cells[i, "A"].Value.ToString(), routeSheet.Cells[i, "B"].Value.ToString(),
                            routeSheet.Cells[i, "C"].Value.ToString(), routeSheet.Cells[i, "D"].Value.ToString());
                        newRoute.Row = i;
                        routes.Add(newRoute);
                        listView.Items.Add(newRoute);
                    }
                }   
            }

            routeBook.Save();
        }

        private void DeleteRoutes(Gui.CheckedListBox listView, Worksheet routeSheet, bool isPreset)
        {
            int count = 2;

            for (int i = listView.Items.Count - 1; i >= 0; i--)
            {
                if (listView.GetItemChecked(i))
                {
                    routes[i].deleteRoute(routeSheet, routeBook, isPreset, reader);
                    routes.RemoveAt(i);
                    listView.Items.Remove(listView.Items[i]);

                    if (!isPreset)
                    {
                        foreach (Route route in routes)
                        {
                            route.Row--;
                        }
                    }
                }
            }

            if (isPreset)
            {
                foreach (Route preset in routes)
                {
                    dataSheet.Cells[count, preset.Column].Value = preset.grade;
                    count++;
                }
            }
            routeBook.Save();
        }

        private void MassSubmit(Gui.RichTextBox gradeBox, Gui.RichTextBox colorBox, Gui.RichTextBox wallBox, Gui.RichTextBox setterBox, Worksheet routeSheet)
        {
            //Find last real row
            int lastUsedRow = routeSheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            int count = lastUsedRow + 1;

            if(gradeBox.Lines.Count() == colorBox.Lines.Count() && gradeBox.Lines.Count() == wallBox.Lines.Count() && gradeBox.Lines.Count() == setterBox.Lines.Count())
            {
                for(int i = 0; i < gradeBox.Lines.Count(); i++)
                {
                    routeSheet.Cells[count, "A"].Value = myTI.ToTitleCase(gradeBox.Lines[i]);
                    routeSheet.Cells[count, "B"].Value = myTI.ToTitleCase(wallBox.Lines[i]);
                    routeSheet.Cells[count, "C"].Value = myTI.ToTitleCase(setterBox.Lines[i]);
                    routeSheet.Cells[count, "D"].Value = myTI.ToTitleCase(colorBox.Lines[i]);

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

            routeBook.Save();
        }

        private void submitParams(Gui.TextBox textbox, string column)
        {
            string cell;

            //Find last real row
            int lastUsedRow = dataSheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            //Create a hashset
            HashSet<string> paramSet;
            List<string> paramList;
            paramList = new List<string>(textbox.Text.Split(','));

            //Add current options to HashSet to avoid duplicates
            for (int i = 2; i <= lastUsedRow; i++)
            {
                cell = dataSheet.Cells[i, column].Value;
                if (cell != null && cell.ToString() != "") paramList.Add(cell);
            }

            dataSheet.Range[column + 2 + ":" + column + (lastUsedRow + 1)].Clear();
            routeBook.Save();

            //Refill column with values
            int count = 2;

            for (int i = 0; i < paramList.Count; i++)
            {
                paramList[i] = paramList[i].Trim();
                paramList[i] = myTI.ToTitleCase(paramList[i]);
            }

            //Convert paramList to set so it can remove duplicates
            paramSet = paramList.ToHashSet();

            

            foreach(string param in paramSet)
            {
                dataSheet.Cells[count, column].Value = param;
                count++;
            }

            routeBook.Save();
        }

        private void receiveGoal(routeOrBoulder dataType, Gui.TabPage page)
        {
            //Fields
            string column;
            int iterate;
            int count = 0;
            List<Gui.TextBox> entryList = new List<Gui.TextBox>();
            bool allValid = true;

            //Set relevant information
            if (dataType == routeOrBoulder.Boulder)
            {
                column = "A";
                iterate = 9;
            }
            else
            {
                column = "B";
                iterate = 7;
            }

            while (count < iterate)
            {
                foreach (Control control in page.Controls)
                {
                    if (control is Gui.TextBox)
                    {
                        entryList.Add((Gui.TextBox)control);
                    }
                }

                entryList.Sort((a, b) => String.Compare(a.Name, b.Name));

                foreach (Gui.TextBox entry in entryList)
                {
                    if (int.TryParse(entry.Text, out int dataInt)) dataSheet.Cells[(count + 2), column].Value = dataInt;
                    else
                    {
                        dataSheet.Cells[(count + 2), column].Value = 0;
                        allValid = false;
                    }
                    entry.Text = "";
                    count++;
                }

            }

            if (allValid) MessageBox.Show("Data Entered!");
            else MessageBox.Show("Data Entered! (Invalid entries were converted to 0)");

            routeBook.Save();
        }

        private void UpdateChart(Worksheet routeSheet, Chart.Chart chart, bool isOnStartup)
        {
            //Fields
            int[] gradeList;
            int[] goalList;
            bool isPlus;

            string replace;
            int gradeMinus;
            string column;
            int gradeInt;
            int overCount = 0;

            bool isBoulder = false;

            //Find last real row
            int lastUsedRow = routeSheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            int dataLastUsedRow = dataSheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            if (routeSheet.Name == "Boulders") 
            {
                gradeList = new int[10];
                goalList = new int[10];
                replace = "V";
                gradeMinus = 0;
                column = "A";
                isBoulder = true;
            } 
            else
            {
                gradeList = new int[12];
                goalList = new int[12];
                replace = "5.";
                gradeMinus = 6;
                column = "B";
            }

            //Get data for current graph
            
            for (int i = 2; i <= lastUsedRow; i++)
            {
                if (routeSheet.Cells[i, "A"].Value != null && routeSheet.Cells[i, "A"].Value.ToString() != "")
                {
                    string tempGrade = routeSheet.Cells[i, "A"].Value.ToString();

                    //Fix broken spreadsheet
                    if(tempGrade == "5.1")
                    {
                        routeSheet.Cells[i, "A"].Value = "'5.10-";
                        gradeList[4]++;
                        break;
                    }

                    if (isBoulder)
                    {
                        gradeInt = int.Parse(tempGrade.Replace(replace, ""));
                        gradeList[gradeInt - gradeMinus]++;
                    }
                    else
                    {
                        try
                        {
                            gradeInt = int.Parse(tempGrade.Replace(replace, "").Replace("+", ""));
                            isPlus = true;
                        }
                        catch (FormatException)
                        {
                            gradeInt = int.Parse(tempGrade.Replace(replace, "").Replace("-", ""));
                            isPlus = false;
                        }

                        if(gradeInt >= 10)
                        {
                            if (isPlus)
                            {
                                gradeList[2 * gradeInt - 16 + 1]++;
                            }
                            else
                            {
                                gradeList[2 * gradeInt - 16]++;
                            }
                        }
                        else
                        {
                            gradeList[gradeInt - gradeMinus]++;
                        }
                    }
                    
                }
                
            }

            //Get data for goal graph
            for (int i = 2; i <= goalList.Length + 1; i++)
            {
                if (dataSheet.Cells[i, column].Value != null && dataSheet.Cells[i, column].Value.ToString() != "")
                {
                    goalList[i-2] = int.Parse(dataSheet.Cells[i, column].Value.ToString());
                }

            }

            //Set initial chart data
            chart.Series[0].Points.Clear();
            chart.Series[1].Points.Clear();

            if (isOnStartup)
            {
                chart.ChartAreas[0].AxisX.Interval = 1;
                chart.ChartAreas[0].AxisY.Interval = 5;

                chart.Series[0].IsValueShownAsLabel = true;
                chart.Series[1].IsValueShownAsLabel = true;
            }

            //Assign data
            if (routeSheet.Name == "Boulders")
            {
                for (int i = 0; i <= 9; i++)
                {
                    chart.Series[0].Points.AddXY("V" + i, gradeList[i]);
                    chart.Series[1].Points.AddXY("V" + i, goalList[i]);
                }
            }

            else for (int i = 0; i <= 11; i++)
                {
                    if(i < 4)
                    {
                        chart.Series[0].Points.AddXY("5." + (i + gradeMinus), gradeList[i]);
                        chart.Series[1].Points.AddXY("5." + (i + gradeMinus), goalList[i]);
                    }
                    else
                    {
                        if(i % 2 != 0)
                        {
                            chart.Series[0].Points.AddXY("5." + (i + gradeMinus - overCount) + "+", gradeList[i]);
                            chart.Series[1].Points.AddXY("5." + (i + gradeMinus - overCount) + "+", goalList[i]);
                        }
                        else
                        {
                            chart.Series[0].Points.AddXY("5." + (i + gradeMinus - overCount) + "-", gradeList[i]);
                            chart.Series[1].Points.AddXY("5." + (i + gradeMinus - overCount) + "-", goalList[i]);
                            overCount++;
                        }
                    }
                }

            if (isBoulder)
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
                ropeGoalBox1.Text = goalList[0].ToString();
                ropeGoalBox2.Text = goalList[1].ToString();
                ropeGoalBox3.Text = goalList[2].ToString();
                ropeGoalBox4.Text = goalList[3].ToString();
                ropeGoalBox5.Text = goalList[4].ToString();
                ropeGoalBox6.Text = goalList[5].ToString();
                ropeGoalBox7.Text = goalList[6].ToString();
                ropeGoalBox8.Text = goalList[7].ToString();
                ropeGoalBox9.Text = goalList[8].ToString();
                ropeGoalBox10.Text = goalList[9].ToString();
                ropeGoalBox11.Text = goalList[10].ToString();
                ropeGoalBox12.Text = goalList[11].ToString();
            }

            routeBook.Save();
            
        }

        private void ResetBoxes(Gui.ComboBox gradeBox, Gui.ComboBox colorBox, Gui.ComboBox wallBox, Gui.ComboBox setterBox)
        {
            gradeBox.SelectedIndex = gradeBox.FindStringExact("Grade:");
            colorBox.SelectedIndex = colorBox.FindStringExact("Color:");
            wallBox.SelectedIndex = wallBox.FindStringExact("Wall:");
            setterBox.SelectedIndex = setterBox.FindStringExact("Setter:");
        }

        private void UpdateOptions()
        {
            List<string> setterOptions = new List<string>();
            List<string> colorOptions = new List<string>();
            List<string> wallOptions = new List<string>();
            List<string> listWallOptions = new List<string>();
            List<string> ropeWallOptions = new List<string>();
            List<string> listRopeWallOptions = new List<string>();

            List<string> listParamOptions = new List<string>()
            {
                "Presets:",
                "Colors",
                "Setters",
                "Boulder Walls",
                "Rope Walls"
            };

            List<string> boulderGradeOptions = new List<string>()
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

            List<string> ropeGradeOptions = new List<string>()
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

            //Find last real row
            int lastUsedRow = dataSheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            setterOptions.Add("Setter:");
            colorOptions.Add("Color:");
            wallOptions.Add("Wall:");
            listWallOptions.Add("Select:");
            ropeWallOptions.Add("Wall:");
            listRopeWallOptions.Add("Select:");


            for (int i = 2; i <= lastUsedRow; i++)
            {
                if (dataSheet.Cells[i, 3].Value != null && dataSheet.Cells[i, 3].Value.ToString() != "") setterOptions.Add(dataSheet.Cells[i, 3].Value.ToString());
                if (dataSheet.Cells[i, 4].Value != null && dataSheet.Cells[i, 4].Value.ToString() != "") colorOptions.Add(dataSheet.Cells[i, 4].Value.ToString());
                if (dataSheet.Cells[i, 5].Value != null && dataSheet.Cells[i, 5].Value.ToString() != "") wallOptions.Add(dataSheet.Cells[i, 5].Value.ToString());
                if (dataSheet.Cells[i, 5].Value != null && dataSheet.Cells[i, 5].Value.ToString() != "") listWallOptions.Add(dataSheet.Cells[i, 5].Value.ToString());
                if (dataSheet.Cells[i, 6].Value != null && dataSheet.Cells[i, 6].Value.ToString() != "") ropeWallOptions.Add(dataSheet.Cells[i, 6].Value.ToString());
                if (dataSheet.Cells[i, 6].Value != null && dataSheet.Cells[i, 6].Value.ToString() != "") listRopeWallOptions.Add(dataSheet.Cells[i, 6].Value.ToString());
            }
            
            
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

        public Form1()
        {
            InitializeComponent();

            if (firstTime) tabControl1.SelectTab(2);
            else tabControl1.SelectTab(0);
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
            Route newRoute = new Route(boulderGradeBox.Text, boulderWallBox.Text, boulderSetterBox.Text, boulderColorBox.Text);
            newRoute.submitRoute(boulderSheet, routeBook, reader);
            ResetBoxes(boulderGradeBox, boulderColorBox, boulderWallBox, boulderSetterBox);
            UpdateChart(boulderSheet, boulderChart, false);
        }

        private void ropeSubmitButton_Click(object sender, EventArgs e)
        {
            Route newRoute = new Route(ropeGradeBox.Text, ropeWallBox.Text, ropeSetterBox.Text, ropeColorBox.Text);

            
            newRoute.submitRoute(ropeSheet, routeBook, reader);

            
            ResetBoxes(ropeGradeBox, ropeColorBox, ropeWallBox, ropeSetterBox);

            
            UpdateChart(ropeSheet, ropeChart, false);
        }

        //Route list
        private void boulderDeleteSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            ListRoutes(boulderDeleteSelect, boulderDeleteBox, boulderSheet, "", false);
        }

        private void boulderDelButton_Click(object sender, EventArgs e)
        {
            
            DeleteRoutes(boulderDeleteBox, boulderSheet, false);

            
            UpdateChart(boulderSheet, boulderChart, false);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            ListRoutes(ropeDeleteSelect, ropeDeleteList, ropeSheet, "", false);
        }

        private void ropeDeleteButton_Click(object sender, EventArgs e)
        {
            
            DeleteRoutes(ropeDeleteList, ropeSheet, false);

            
            UpdateChart(ropeSheet, ropeChart, false);
        }

        private void presetBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string column = "";

            switch (presetBox.Text)
            {
                case "Setters":
                    column = "C";
                    break;
                case "Colors":
                    column = "D";
                    break;
                case "Boulder Walls":
                    column = "E";
                    break;
                case "Rope Walls":
                    column = "F";
                    break;
            }

            
            if(column != "") ListRoutes(presetBox, presetListBox, dataSheet, column, true);
        }

        private void presetDeleteButton_Click(object sender, EventArgs e)
        {
            
            DeleteRoutes(presetListBox, dataSheet, true);
        }

        private void boulderCurveSubmitButton_Click(object sender, EventArgs e)
        {
            
            receiveGoal(routeOrBoulder.Boulder, curveControl.SelectedTab);
            
            UpdateChart(boulderSheet, boulderChart, false);
        }

        private void ropeCurveSubmit_Click(object sender, EventArgs e)
        {
            receiveGoal(routeOrBoulder.Route, curveControl.SelectedTab);
            
            UpdateChart(ropeSheet, ropeChart, false);
        }


        private void boulderMassSubmitButton_Click(object sender, EventArgs e)
        {
            MassSubmit(boulderGradeMassSubmit, boulderColorMassSubmit, boulderWallMassSubmit, boulderSetterMassSubmit, boulderSheet);

            
            UpdateChart(ropeSheet, ropeChart, false);
        }

        private void ropeMassSubmitButton_Click(object sender, EventArgs e)
        {
            MassSubmit(ropeGradeMassSubmit, ropeColorMassSubmit, ropeWallMassSubmit, ropeSetterMassSubmit, ropeSheet);

            UpdateChart(ropeSheet, ropeChart, false);
        }

        private void parameterSubmitButton_Click(object sender, EventArgs e)
        {
            if(setterOptionsBox.Text != "") submitParams(setterOptionsBox, "C");

            if(colorOptionsBox.Text !="") submitParams(colorOptionsBox, "D");

            if(wallOptionsBox.Text != "") submitParams(wallOptionsBox, "E");

            if(rWallOptionsBox.Text != "") submitParams(rWallOptionsBox, "F");

            UpdateOptions();

            setterOptionsBox.Clear();
            colorOptionsBox.Clear();
            wallOptionsBox.Clear();
            rWallOptionsBox.Clear();

            Gui.MessageBox.Show("Options Submitted!");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            routeBook.Close(0);
            routeBook = null;
            reader.Quit();
            reader = null;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CheckForOpenFile();
        }
    }
}
