using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;

namespace WPF_EEOI_Calculator_v2
{
    [Serializable]
    [DataContract]
    public abstract class Observable : INotifyPropertyChanged
    {
        #region Interfaces
        [field: NonSerializedAttribute()]
        public event PropertyChangedEventHandler PropertyChanged;

        public void NotifyChange([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion
    }
}
