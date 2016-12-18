using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ObjectStorageTests.Types
{
    public class StoredTypeNotifyPropertyChanged : INotifyPropertyChanged, IStoredType
    {
        private string _Value1;
        private int _Value2;

        public string Value1
        {
            get { return _Value1; }
            set { SetField(ref _Value1, value); }
        }
        public int Value2
        {
            get { return _Value2; }
            set { SetField(ref _Value2, value); }
        }

        public StoredTypeNotifyPropertyChanged(string Value1, int Value2)
        {
            this.Value1 = Value1;
            this.Value2 = Value2;
        }

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        protected bool SetField<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        #endregion

    }
}
