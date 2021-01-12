using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace onSet
{
    static class Program
    {
        [STAThread]        
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);


            //Create view
            Form1 view = new Form1();
            DBManagement data = new DBManagement();
            Controller controller = new Controller(data, view);
            //Set form controller
            view.SetController(controller);
            //Subscribe to events
            data.NewGraphData += view.ExcelManagement_NewGraphData;
            data.PresetsUpdated += view.ExcelManagement_PresetsUpdated;
            data.NewListView += view.ExcelManagement_NewListView;
            //Initial program setup
            view.Setup();

            Application.Run(view);
        }
    }
}
