using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace onSet
{
    class PresetsUpdatedEventArgs : EventArgs
    {
        public List<string> setterOptions;
        public List<string> colorOptions;
        public List<string> wallOptions;
        public List<string> listWallOptions;
        public List<string> ropeWallOptions;
        public List<string> listRopeWallOptions;
        public List<string> listParamOptions;
        public List<string> boulderGradeOptions;
        public List<string> ropeGradeOptions;
    }
}
