using System.Windows.Forms;
using Microsoft.Xrm.Sdk;
using McTools.Xrm.Connection;

namespace XrmToolBox.Extensibility
{
    public class PluginControlBase : UserControl, XrmToolBox.Extensibility.Interfaces.IXrmToolBoxPluginControl
    {
        public IOrganizationService Service { get; set; }
        public ConnectionDetail ConnectionDetail { get; set; }

        // Minimal no-op shims so plugin code can compile
        public void ExecuteMethod(object info) { }
        public void WorkAsync(WorkAsyncInfo info) { }
    }
}
