namespace SimphonyDataEntryHelper
{
    public class ModeViewModel : ViewModel
    {
        private Mode _mode;
        public Mode Mode
        {
            get => _mode;
            set
            {
                if (_mode == value)
                    return;
                _mode = value;
                OnPropertyChanged(nameof(Mode));
            }
        }

        private string _name;
        public string Name
        {
            get => _name;
            set
            {
                if (_name == value)
                    return;
                _name = value;
                OnPropertyChanged(nameof(Name));
            }
        }
    }
}
