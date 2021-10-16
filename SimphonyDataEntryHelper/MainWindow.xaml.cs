using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using TextCopy;
using WindowsInput;
using Code = WindowsInput.Native.VirtualKeyCode;

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

    public class ViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(string property) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
    }

    public enum Mode
    {
        MasterEntry,
        CondimentEntry,
        PriceEntry
    }

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

    public class MainViewModel : ViewModel
    {
        private readonly InputSimulator _sim;

        private ModeViewModel _selectedMode;
        public ModeViewModel SelectedMode
        {
            get => _selectedMode;
            set
            {
                if (_selectedMode == value)
                    return;
                _selectedMode = value;
                OnPropertyChanged(nameof(SelectedMode));
            }
        }

        private ObservableCollection<ModeViewModel> _modes;
        public ObservableCollection<ModeViewModel> Modes
        {
            get => _modes;
            set
            {
                if (_modes == value)
                    return;
                _modes = value;
                OnPropertyChanged(nameof(Modes));
            }
        }

        private Task ActiveTask
        {
            get
            {
                return SelectedMode.Mode switch
                {
                    Mode.MasterEntry => MenuItemEntry(),
                    Mode.CondimentEntry => CondimentEntry(),
                    _ => PriceEntry()
                };
            }
        }

        public MainViewModel(InputSimulator sim)
        {
            _sim = sim;
            Modes = new ObservableCollection<ModeViewModel>()
            {
                new ModeViewModel(){ Mode = Mode.CondimentEntry, Name = "Condiment Entry"},
                new ModeViewModel(){ Mode = Mode.MasterEntry, Name = "Master Entry"},
                new ModeViewModel(){ Mode = Mode.PriceEntry, Name = "Price Entry"}
            };
            Task.Factory.StartNew(Run);
        }

        public async Task Run()
        {
            while (true)
            {
                await Task.Delay(50);
                if (isKeyDown(Code.OEM_PERIOD) && isKeyDown(Code.CONTROL))
                {
                    await ActiveTask;
                }
            }
        }

        private async Task MenuItemEntry()
        {
            var cbText = await ClipboardService.GetTextAsync();
            var items = cbText.Split(Environment.NewLine);
            foreach (var item in items)
            {
                await ClipboardService.SetTextAsync(cap(item).Trim());
                await Task.Delay(1300);
                keydown(Code.CONTROL);
                keypress(Code.VK_V);
                keyup(Code.CONTROL);
                await Task.Delay(1300);
                lmouse();
                await Task.Delay(2500);
                keypress(Code.SPACE);
            }
        }

        private async Task CondimentEntry()
        {
            var cbText = await ClipboardService.GetTextAsync();
            var items = cbText.Split(Environment.NewLine);
            foreach (var item in items)
            {
                await ClipboardService.SetTextAsync(low(item).Trim());
                await Task.Delay(1300);
                keydown(Code.CONTROL);
                keypress(Code.VK_V);
                keyup(Code.CONTROL);
                await Task.Delay(1300);
                lmouse();
                await Task.Delay(2500);
                keypress(Code.SPACE);
            }
        }

        private async Task PriceEntry()
        {
            var cbText = await ClipboardService.GetTextAsync();
            var items = cbText.Split(Environment.NewLine);
            foreach (var item in items)
            {
                await ClipboardService.SetTextAsync(item.Trim());
                await Task.Delay(150);
                keydown(Code.CONTROL);
                keypress(Code.VK_A);
                keyup(Code.CONTROL);
                foreach (var chr in item.Trim())
                {
                    keypress(chr switch
                    {
                        '0' => Code.VK_0,
                        '1' => Code.VK_1,
                        '2' => Code.VK_2,
                        '3' => Code.VK_3,
                        '4' => Code.VK_4,
                        '5' => Code.VK_5,
                        '6' => Code.VK_6,
                        '7' => Code.VK_7,
                        '8' => Code.VK_8,
                        '9' => Code.VK_9,
                        _ => Code.OEM_PERIOD
                    });
                }
                await Task.Delay(100);
                keypress(Code.RETURN);
            }
        }

        private string cap(string item)
        {
            bool capitalize = true;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < item.Length; i++)
            {
                if (capitalize)
                {
                    sb.Append(item[i].ToString().ToUpper());
                    capitalize = false;
                }
                else
                    sb.Append(item[i].ToString().ToLower());

                if (item[i] == ' ')
                    capitalize = true;
            }
            return sb.ToString();
        }
        public string low(string item) => item.ToLower();
        private bool isKeyDown(Code code) => _sim.InputDeviceState.IsHardwareKeyDown(code);
        private void keydown(Code code) => _sim.Keyboard.KeyDown(code);
        private void keyup(Code code) => _sim.Keyboard.KeyUp(code);
        private void keypress(Code code) => _sim.Keyboard.KeyPress(code);
        private void lmouse() => _sim.Mouse.LeftButtonClick();
    }
}
