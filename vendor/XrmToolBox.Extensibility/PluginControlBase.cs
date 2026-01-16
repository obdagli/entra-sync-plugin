using System.Windows.Forms;
using Microsoft.Xrm.Sdk;
using McTools.Xrm.Connection;

namespace XrmToolBox.Extensibility
{
    public class PluginControlBase : UserControl
    {
        public IOrganizationService Service { get; set; }
        public ConnectionDetail ConnectionDetail { get; set; }

        // Minimal no-op shim so plugin code can call ExecuteMethod
        public void ExecuteMethod(object info) { }
    }
}