using System;
using XrmToolBox.Extensibility;

namespace BulkAssociate
{
    public partial class BulkAssociate : PluginControlBase
    {
        public BulkAssociate()
        {
            InitializeComponent();
        }

        private void tsbClose_Click(object sender, EventArgs e)
        {
            CloseTool();
        }
    }
}
