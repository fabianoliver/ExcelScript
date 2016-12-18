using System;

namespace RoslynScripting.Internal
{
    internal class DelegateDisposable : IDisposable
    {
        private readonly Action OnDisposing;
        private bool m_disposed = false;

        public DelegateDisposable(Action OnDisposing)
        {
            if (OnDisposing == null)
                throw new ArgumentNullException(nameof(OnDisposing));

            this.OnDisposing = OnDisposing;
        }

        public void Dispose()
        {
            if (!m_disposed)
            {
                OnDisposing();
                m_disposed = true;
            }
        }
    }
}
