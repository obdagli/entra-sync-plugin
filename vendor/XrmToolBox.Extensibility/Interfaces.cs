namespace XrmToolBox.Extensibility.Interfaces
{
    public interface IPlugin { }

    public interface IXrmToolBoxPluginControl { }

    public interface IXrmToolBoxPlugin : IPlugin
    {
        IXrmToolBoxPluginControl GetControl();
    }
}
