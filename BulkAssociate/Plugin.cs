using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace BulkAssociate
{
    [Export(typeof(IXrmToolBoxPlugin)),
    ExportMetadata("BackgroundColor", "MediumBlue"),
    ExportMetadata("PrimaryFontColor", "White"),
    ExportMetadata("SecondaryFontColor", "LightGray"),
    ExportMetadata("SmallImageBase64", ""),
    ExportMetadata("BigImageBase64", ""),
    ExportMetadata("Name", "Bulk Associate"),
    ExportMetadata("Author", "Remon Kamel"),
    ExportMetadata("Description", "Associate in Bulk")]
    public class Plugin : PluginBase
    {
        public override IXrmToolBoxPluginControl GetControl()
        {
            return new BulkAssociate();
        }
    }
}
