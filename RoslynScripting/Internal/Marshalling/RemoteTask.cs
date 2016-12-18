using System;
using System.Threading;
using System.Threading.Tasks;

namespace RoslynScripting.Internal.Marshalling
{
    /// <summary>
    /// This code is from http://stackoverflow.com/questions/15142507/deadlock-when-combining-app-domain-remoting-and-tasks
    /// </summary>
    internal static class RemoteTask
    {
        public static async Task<T> ClientComplete<T>(RemoteTask<T> remoteTask, CancellationToken cancellationToken)
        {
            T result;

            using (cancellationToken.Register(remoteTask.Cancel))
            {
                RemoteTaskCompletionSource<T> tcs = new RemoteTaskCompletionSource<T>();
                remoteTask.Complete(tcs);
                result = await tcs.Task;
            }

            await Task.Yield();

            return result;
        }

        public static RemoteTask<T> ServerStart<T>(Func<CancellationToken, Task<T>> func)
        {
            return new RemoteTask<T>(func);
        }
    }

    internal class RemoteTask<T> : MarshalByRefObject
    {
        readonly CancellationTokenSource cts = new CancellationTokenSource();
        readonly Task<T> task;

        internal RemoteTask(Func<CancellationToken, Task<T>> starter)
        {
            this.task = starter(cts.Token);
        }

        internal void Complete(RemoteTaskCompletionSource<T> tcs)
        {
            task.ContinueWith(t =>
            {
                if (t.IsFaulted)
                {
                    tcs.TrySetException(t.Exception);
                }
                else if (t.IsCanceled)
                {
                    tcs.TrySetCancelled();
                }
                else
                {
                    tcs.TrySetResult(t.Result);
                }
            }, TaskContinuationOptions.ExecuteSynchronously);
        }

        internal void Cancel()
        {
            cts.Cancel();
        }
    }

}
