using System;
using System.ComponentModel;

namespace XrmToolBox.Extensibility
{
    public class WorkAsyncCallbackArgs
    {
        public Exception Error { get; set; }
        public object Result { get; set; }
    }

    public class WorkAsyncInfo
    {
        public string Message { get; set; }
        public Action<BackgroundWorker, DoWorkEventArgs> Work { get; set; }
        public Action<WorkAsyncCallbackArgs> PostWorkCallBack { get; set; }
        public Action<ProgressChangedEventArgs> ProgressChanged { get; set; }
        public object Result { get; set; }
    }
}
