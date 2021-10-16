using WindowsInput;
using System.Windows;

namespace SimphonyDataEntryHelper
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel(new InputSimulator());
        }

    }
}
