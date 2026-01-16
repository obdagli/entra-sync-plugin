using System.ComponentModel.Composition;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace EntraSyncPlugin
{
    [Export(typeof(IXrmToolBoxPlugin))]
    [ExportMetadata("Name", "Entra Sync Plugin")]
    [ExportMetadata("Description", "Sync Entra user data using impersonation in Dataverse.")]
    [ExportMetadata("SmallImageBase64", null)]
    [ExportMetadata("BigImageBase64", null)]
    [ExportMetadata("BackgroundColor", "LightSteelBlue")]
    [ExportMetadata("PrimaryFontColor", "Black")]
    [ExportMetadata("SecondaryFontColor", "DimGray")]
    public class EntraSyncPluginEntry : PluginBase
    {
        public override IXrmToolBoxPluginControl GetControl()
        {
            return new EntraSyncControl();
        }
    }
}
