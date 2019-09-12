using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DWGViewer
{
    public partial class DWGViewer : UserControl
    {
        public string dwgFilePath { get; set; }
        public DWGViewer()
        {
            InitializeComponent();
            
        }

        public void loadFile(string filePath)
        {
            dwgFilePath = filePath;
            ACADViewer.PutSourcePath(dwgFilePath);
        }

    }
}
