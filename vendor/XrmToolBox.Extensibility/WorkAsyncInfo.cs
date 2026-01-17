using System;
using System.ComponentModel;

namespace XrmToolBox.Extensibility
{
    public class WorkAsyncInfo
    {
        public string Message { get; set; }
        public Action<BackgroundWorker, DoWorkEventArgs> Work { get; set; }
        public Action<RunWorkerCompletedEventArgs> PostWorkCallBack { get; set; }
        public Action<ProgressChangedEventArgs> ProgressChanged { get; set; }
        public object Result { get; set; }
    }
}
