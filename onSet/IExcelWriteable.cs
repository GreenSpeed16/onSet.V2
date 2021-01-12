using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace onSet
{
    public interface IExcelWriteable
    {
        string PrimaryKey { get; set; }
        void Submit();
    }
}
