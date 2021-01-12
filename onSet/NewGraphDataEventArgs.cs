using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Charting = System.Windows.Forms.DataVisualization.Charting;

namespace onSet
{
    class NewGraphDataEventArgs : EventArgs
    {
        public List<int> currentData;
        public List<int> goalData;
        public List<string> gradeLabels;
        public Charting.Chart chart;
    }
}
