using System;
using System.Threading.Tasks;

namespace RoslynScripting.Internal.Marshalling
{
    /// <summary>
    /// This code is from http://stackoverflow.com/questions/15142507/deadlock-when-combining-app-domain-remoting-and-tasks
    /// </summary>
    internal class RemoteTaskCompletionSource<T> : MarshalByRefObject
    {
        readonly TaskCompletionSource<T> tcs = new TaskCompletionSource<T>();

        public bool TrySetResult(T result) { return tcs.TrySetResult(result); }
        public bool TrySetCancelled() { return tcs.TrySetCanceled(); }
        public bool TrySetException(Exception ex) { return tcs.TrySetException(ex); }

        public Task<T> Task
        {
            get
            {
                return tcs.Task;
            }
        }
    }
}
