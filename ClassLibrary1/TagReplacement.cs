using System.ComponentModel;

namespace ClassLibrary1
{
    public class TagReplacement : INotifyPropertyChanged
    {
        public TagReplacement(string tag, string value = "")
        {
            _tag = tag;
            _value = value;
        }

        private string _tag;
        public string Tag
        {
            get => _tag;
            set
            {
                if (_tag == value) return;
                _tag = value;
                OnPropertyChanged("State");
            }
        }

        private string _value;
        public string? Value
        {
            get => _value;
            set
            {
                if (_value == value) return;
                _value = value;
                OnPropertyChanged("State");
            }
        }


        public event PropertyChangedEventHandler? PropertyChanged;
        public void OnPropertyChanged(string info) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(info));
    }
}
