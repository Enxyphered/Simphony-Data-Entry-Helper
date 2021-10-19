using System;
using System.Text;
using System.Linq;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using TextCopy;
using WindowsInput;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using FlaUI.Core.Definitions;
using Code = WindowsInput.Native.VirtualKeyCode;
using FlaUI.Core.Conditions;
using System.Threading;
using FlaUI.Core.Input;
using FlaUI.Core;

namespace SimphonyDataEntryHelper
{
    public class MenuItemEntryConfigViewModel : ViewModel
    {
        private int _delayBeforePaste;
        public int DelayBeforePaste
        {
            get => _delayBeforePaste;
            set
            {
                if (_delayBeforePaste == value)
                    return;
                _delayBeforePaste = value;
                OnPropertyChanged(nameof(DelayBeforePaste));
            }
        }

        private int _delayBeforeOk;
        public int DelayBeforeOk
        {
            get => _delayBeforeOk;
            set
            {
                if (_delayBeforeOk == value)
                    return;
                _delayBeforeOk = value;
                OnPropertyChanged(nameof(DelayBeforeOk));
            }
        }

        private int _delayBeforeAddAnotherItem;
        public int DelayBeforeAddAnotherItem
        {
            get => _delayBeforeAddAnotherItem;
            set
            {
                if (_delayBeforeAddAnotherItem == value)
                    return;
                _delayBeforeAddAnotherItem = value;
                OnPropertyChanged(nameof(DelayBeforeAddAnotherItem));
            }
        }
    }

    public class MainViewModel : ViewModel
    {
        private readonly InputSimulator _sim;
        private CancellationTokenSource _cSouce;
        private Application _app;

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

        private CancellationToken token { get => _cSouce.Token; }

        public MainViewModel(InputSimulator sim)
        {
            _sim = sim;
            _cSouce = new CancellationTokenSource();
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
            var process = Process.GetProcessesByName("EMC")[0];
            _app = Application.Attach(process);
            while (true)
            {
                await Task.Delay(50);
                if (isKeyDown(Code.OEM_PERIOD) && isKeyDown(Code.CONTROL))
                {
                    _cSouce = new CancellationTokenSource();
                    await ActiveTask;
                    await Task.Delay(100);
                }
                if (isKeyDown(Code.OEM_COMMA) && isKeyDown(Code.CONTROL))
                {
                    _cSouce.Cancel();
                }
            }
        }

        private async Task MenuItemEntry()
        {
            using (var automation = new UIA3Automation())
            {
                var cbText = await ClipboardService.GetTextAsync();
                var items = cbText.Split('\n');
                foreach (var item in items)
                {
                    if (token.IsCancellationRequested)
                        break;

                    var window = _app.GetMainWindow(automation);
                    var addMenuItemsWin = window.FindFirstChild(cf => cf.ByName("Add Menu Items"));
                    var wiz = addMenuItemsWin.FindFirstChild(cf => cf.ByAutomationId("MasterWizard"));
                    var addWithTemplate = wiz.FindFirstChild(cf => cf.ByName("Add Menu Item Master With Template"));
                    var nameTextBox = addWithTemplate.FindFirstChild("textBoxName");
                    var mainWizOk = wiz.FindFirstChild(cf => cf.ByName("OK"));
                    nameTextBox.AsTextBox().Text = cap(item).Trim();
                    mainWizOk.AsButton().Click();
                    await Task.Delay(1300);
                    var dialog = await element(automation.GetDesktop(), cf => cf.ByName("Item Added Successfully"));
                    var dialogYes = await element(dialog, cf => cf.ByName("Yes"));
                    dialogYes.AsButton().Click();
                }
            }
        }

        private async Task CondimentEntry()
        {
            using (var automation = new UIA3Automation())
            {
                var cbText = await ClipboardService.GetTextAsync();
                var items = cbText.Split('\n');
                foreach (var item in items)
                {
                    if (token.IsCancellationRequested)
                        break;

                    var window = _app.GetMainWindow(automation);
                    var addMenuItemsWin = window.FindFirstChild(cf => cf.ByName("Add Menu Items"));
                    var wiz = addMenuItemsWin.FindFirstChild(cf => cf.ByAutomationId("MasterWizard"));
                    var addWithTemplate = wiz.FindFirstChild(cf => cf.ByName("Add Menu Item Master With Template"));
                    var nameTextBox = addWithTemplate.FindFirstChild("textBoxName");
                    var mainWizOk = wiz.FindFirstChild(cf => cf.ByName("OK"));
                    nameTextBox.AsTextBox().Text = low(item).Trim();
                    mainWizOk.AsButton().Click();
                    await Task.Delay(1300);
                    var dialog = await element(automation.GetDesktop(), cf => cf.ByName("Item Added Successfully"));
                    var dialogYes = await element(dialog, cf => cf.ByName("Yes"));
                    dialogYes.AsButton().Click();
                }
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

        private async Task<AutomationElement> element(AutomationElement element, Func<ConditionFactory, ConditionBase> func)
        {
            AutomationElement output = null;
            while (output is null)
            {
                output = element.FindFirstChild(func);
                await Task.Delay(100);
            }
            return output;
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
    }
}
