using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Gui = System.Windows.Forms;

namespace onSet
{
    public class NewListViewEventArgs : EventArgs
    {
        public List<IExcelWriteable> DataList;
        public Gui.CheckedListBox ListView;
    }
}
