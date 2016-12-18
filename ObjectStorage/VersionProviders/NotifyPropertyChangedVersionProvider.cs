using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ObjectStorage.VersionProviders
{
    internal class NotifyPropertyChangedVersionProvider : IVersionProvider, IDisposable
    {
        private WeakReference<INotifyPropertyChanged> m_EventSource;
        private bool m_isDirty = false;
        private int m_Version;

        public int GetVersionFor(object obj)
        {
            if(m_EventSource != null)
            {
                INotifyPropertyChanged oldSource;
                if(m_EventSource.TryGetTarget(out oldSource) && oldSource == obj)
                {
                    if (m_isDirty)
                    {
                        m_isDirty = false;
                        this.m_Version += 1;
                    }

                    return this.m_Version;
                }
            }

            TrySubscribeTo(obj);
            this.m_Version += 1;
            return this.m_Version;
        }

        public void Dispose()
        {
            TryUnsubscribe();
        }

        public static bool CanHandle(Type t)
        {
            return typeof(INotifyPropertyChanged).IsAssignableFrom(t);
        }

        public NotifyPropertyChangedVersionProvider(object InitialValue, int InitialVersion = 1)
        {
            TrySubscribeTo(InitialValue);
            this.m_Version = InitialVersion;
        }

        private void TrySubscribeTo(object obj)
        {
            TryUnsubscribe();

            INotifyPropertyChanged _obj = obj as INotifyPropertyChanged;
            if(_obj != null)
            {
                _obj.PropertyChanged += EventSourcePropertyChanged;
                this.m_EventSource = new WeakReference<INotifyPropertyChanged>(_obj);
            }
        }

        private void EventSourcePropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            m_isDirty = true;
        }

        private void TryUnsubscribe()
        {
            INotifyPropertyChanged oldEventSource;
            if (m_EventSource != null && m_EventSource.TryGetTarget(out oldEventSource))
            {
                oldEventSource.PropertyChanged -= EventSourcePropertyChanged;
                oldEventSource = null;
            }
        }
    }
}
