using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace WpfApplication5
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //globals
        List<string> list = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
            asyncLCFS();
            dropdown_reset.Items.Add("Clock Speeds");
            dropdown_reset.Items.Add("Voltages");
            dropdown_reset.Items.Add("All");
            MessageBox.Show("As always, be careful when making any modifications to your cards (I assume zero responsibility for your card catching fire or starting the first robot vs human war) and reboot after making final modifications to see true results.", "Welcome to ecs87's AMD GPU Control Tool");
        }
        private async void asyncLCFS()
        {
            await Task.Run(() => loopCurrentFanSpeed());
        }
        private void loopCurrentFanSpeed()
        {
            while (true)
            {
                try
                {
                    string firstcheck = "";
                    this.Dispatcher.Invoke(() =>
                    {
                        firstcheck = deviceComboBox.Text;
                    });
                    if (firstcheck == "") { }
                    else
                    {
                        //get bus number
                        string confirmBusNumber = firstcheck;
                        var confirmBusNumber2 = confirmBusNumber.IndexOf("Bus ID: ");
                        var confirmBusNumberPre = confirmBusNumber.Substring(confirmBusNumber2 + 8, 4);
                        var confirmBusNumberFinal = Regex.Replace(confirmBusNumberPre, "[^0-9.]", "");
                        //start
                        Process p2 = new Process();
                        p2.StartInfo.UseShellExecute = false;
                        p2.StartInfo.RedirectStandardOutput = true;
                        p2.StartInfo.CreateNoWindow = true;
                        p2.StartInfo.FileName = @"ecs87NextGearOD.exe";
                        p2.StartInfo.Arguments = "c " + confirmBusNumberFinal;
                        p2.Start();
                        Thread.Sleep(500);
                        //Current Fan Speed sub-substrings...
                        string standard_output = "";
                        while ((standard_output = p2.StandardOutput.ReadLine()) != null)
                        {
                            if (standard_output.Contains("iCurrentFanSpeed : "))
                            {
                                break;
                            }
                        }
                        p2.Close();
                        var indexofCurrentFanSpeed = standard_output.IndexOf("iCurrentFanSpeed : ");
                        var SubstringofCurrentFanSpeed = standard_output.Substring(indexofCurrentFanSpeed + 19);
                        var SubstringofCurrentFanSpeedFinal = Regex.Replace(SubstringofCurrentFanSpeed, "[^0-9.]", "");
                        this.Dispatcher.Invoke(() =>
                        {
                            current_FanSpeed.Text = SubstringofCurrentFanSpeedFinal;
                        });
                    }
                    Thread.Sleep(1000);
                }
                catch { }
            }
        }
        private void button_refreshDeviceCombobox_Click(object sender, RoutedEventArgs e)
        {
            list = new List<string>();
            deviceComboBox.Items.Clear();
            string standard_output;
            Process p = new Process();
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.CreateNoWindow = true;
            p.StartInfo.FileName = @"ecs87NextGearOD.exe";
            p.StartInfo.Arguments = "g";
            /*
            p.OutputDataReceived += (s, f) => Console.WriteLine(f.Data);
            p.Start();
            p.BeginOutputReadLine();
            */
            p.Start();
            while ((standard_output = p.StandardOutput.ReadLine()) != null)
            {
                list.Add(standard_output);
            }

            foreach (var listitem in list)
            {
                if (listitem.Contains("Bus ID"))
                {
                    deviceComboBox.Items.Add(listitem);
                }
            }
            p.Close();
            Process p2 = new Process();
            p2.StartInfo.UseShellExecute = false;
            p2.StartInfo.RedirectStandardOutput = true;
            p2.StartInfo.FileName = @"ecs87NextGearOD.exe";
            p2.StartInfo.Arguments = "f";
            p2.StartInfo.CreateNoWindow = true;
            p2.Start();
            while ((standard_output = p2.StandardOutput.ReadLine()) != null)
            {
                list.Add(standard_output);
            }
            p2.Close();
        }
        private void comboBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string CurrentItem = (string)deviceComboBox.SelectedItem;
            var result = string.Join(" ", list.ToArray());
            var stringList2 = Regex.Split(result, "Bus ID:").ToList();
            foreach (var stringlistitem in stringList2)
            {
                var stringlistitemfinal = " Bus ID:" + stringlistitem;
                try
                {
                    if (stringlistitemfinal.Contains(CurrentItem) && stringlistitemfinal.Contains("Assigned Target Temperature :"))
                    {
                        fansAndTemps(stringlistitem);
                    }
                    if (stringlistitemfinal.Contains(CurrentItem))
                    {
                        loopClocksAndVoltages(stringlistitem, 0);
                        loopClocksAndVoltages(stringlistitem, 1);
                        loopClocksAndVoltages(stringlistitem, 2);
                        loopClocksAndVoltages(stringlistitem, 3);
                        loopClocksAndVoltages(stringlistitem, 4);
                        loopClocksAndVoltages(stringlistitem, 5);
                        loopClocksAndVoltages(stringlistitem, 6);
                        loopClocksAndVoltages(stringlistitem, 7);
                    }
                }
                catch { }
            }
        }
        private void fansAndTemps(string stringlistitem)
        {
            //target temp
            var stringListTTemp = stringlistitem.IndexOf("Assigned Target Temperature :");
            string stringListItemSubStrTTempFinal = "";
            var stringListItemSubStrTTemp = stringlistitem.Substring(stringListTTemp + 30, 3);
            stringListItemSubStrTTempFinal = Regex.Replace(stringListItemSubStrTTemp, "[^0-9.]", "");
            target_Temperature.Text = stringListItemSubStrTTempFinal;
            //min fan speed
            stringListTTemp = stringlistitem.IndexOf("Assigned Minimum/Target fan Limit : ");
            stringListItemSubStrTTempFinal = "";
            stringListItemSubStrTTemp = stringlistitem.Substring(stringListTTemp + 35, 5);
            stringListItemSubStrTTempFinal = Regex.Replace(stringListItemSubStrTTemp, "[^0-9.]", "");
            min_FanSpeed.Text = stringListItemSubStrTTempFinal;
            //max fan speed
            stringListTTemp = stringlistitem.IndexOf("Assigned Minimum/Target fan Limit : ");
            stringListItemSubStrTTempFinal = "";
            stringListItemSubStrTTemp = stringlistitem.Substring(stringListTTemp + 42, 5);
            stringListItemSubStrTTempFinal = Regex.Replace(stringListItemSubStrTTemp, "[^0-9.]", "");
            target_FanSpeed.Text = stringListItemSubStrTTempFinal;
        }
        private void loopClocksAndVoltages(string stringlistitem, int powerLevel)
        {
            int stringList;
            int stringListMem;
            int stringListCV;
            int stringListMV;
            int stringListTTemp;

            stringList = stringlistitem.IndexOf("CORE: Level " + powerLevel + " / ");
            stringListMem = stringlistitem.IndexOf("MEMORY: Level " + powerLevel + " / ");
            stringListTTemp = stringlistitem.IndexOf("Assigned Target Temperature :");

            //Core clock substrings
            var stringListItemSubStr = stringlistitem.Substring(stringList + 16, 4);
            var stringListItemSubStrMem = stringlistitem.Substring(stringListMem + 18, 6);

            //Core voltage sub-substrings...
            var stringListItemSubStrCVPre = stringlistitem.Substring(stringList);
            stringListCV = stringListItemSubStrCVPre.IndexOf("Enabled: ");
            var stringListItemSubStrCV = stringListItemSubStrCVPre.Substring(stringListCV + 13, 4);
            var stringListItemSubStrCVFinal = Regex.Replace(stringListItemSubStrCV, "[^0-9.]", "");

            //Mem voltage sub-substrings...
            var stringListItemSubStrMVPre = stringlistitem.Substring(stringListMem);
            stringListMV = stringListItemSubStrMVPre.IndexOf("Enabled: ");
            var stringListItemSubStrMV = stringListItemSubStrMVPre.Substring(stringListMV + 13, 4);
            var stringListItemSubStrMVFinal = Regex.Replace(stringListItemSubStrMV, "[^0-9.]", "");
            if (stringListItemSubStrMVFinal == "") { stringListItemSubStrMVFinal = "0"; }

            //Mem Clock substrings
            var stringListItemSubStrFinal = Regex.Replace(stringListItemSubStr, "[^0-9.]", "");
            var stringListItemSubStrFinalMem = Regex.Replace(stringListItemSubStrMem, "[^0-9.]", "");

            if (powerLevel == 0)
            {
                PPlayCC1.Text = stringListItemSubStrFinal;
                PPlayMC1.Text = stringListItemSubStrFinalMem;
                PPlayCV1.Text = stringListItemSubStrCVFinal;
                PPlayMV1.Text = stringListItemSubStrMVFinal;
            }
            if (powerLevel == 1)
            {
                PPlayCC2.Text = stringListItemSubStrFinal;
                PPlayMC2.Text = stringListItemSubStrFinalMem;
                PPlayCV2.Text = stringListItemSubStrCVFinal;
                PPlayMV2.Text = stringListItemSubStrMVFinal;
            }
            if (powerLevel == 2)
            {
                PPlayCC3.Text = stringListItemSubStrFinal;
                PPlayMC3.Text = stringListItemSubStrFinalMem;
                PPlayCV3.Text = stringListItemSubStrCVFinal;
                PPlayMV3.Text = stringListItemSubStrMVFinal;
            }
            if (powerLevel == 3)
            {
                PPlayCC4.Text = stringListItemSubStrFinal;
                PPlayMC4.Text = stringListItemSubStrFinalMem;
                PPlayCV4.Text = stringListItemSubStrCVFinal;
                PPlayMV4.Text = stringListItemSubStrMVFinal;
            }
            if (powerLevel == 4)
            {
                PPlayCC5.Text = stringListItemSubStrFinal;
                PPlayMC5.Text = stringListItemSubStrFinalMem;
                PPlayCV5.Text = stringListItemSubStrCVFinal;
                PPlayMV5.Text = stringListItemSubStrMVFinal;
            }
            if (powerLevel == 5)
            {
                PPlayCC6.Text = stringListItemSubStrFinal;
                PPlayMC6.Text = stringListItemSubStrFinalMem;
                PPlayCV6.Text = stringListItemSubStrCVFinal;
                PPlayMV6.Text = stringListItemSubStrMVFinal;
            }
            if (powerLevel == 6)
            {
                PPlayCC7.Text = stringListItemSubStrFinal;
                PPlayMC7.Text = stringListItemSubStrFinalMem;
                PPlayCV7.Text = stringListItemSubStrCVFinal;
                PPlayMV7.Text = stringListItemSubStrMVFinal;
            }
            if (powerLevel == 7)
            {
                PPlayCC8.Text = stringListItemSubStrFinal;
                PPlayMC8.Text = stringListItemSubStrFinalMem;
                PPlayCV8.Text = stringListItemSubStrCVFinal;
                PPlayMV8.Text = stringListItemSubStrFinalMem;
            }
        }
        private void buttonBIOSinfo_Click(object sender, RoutedEventArgs e)
        {
            //get bus number
            string confirmBusNumber = deviceComboBox.Text;
            if (confirmBusNumber == "") { return; }
            string standard_output = "";
            string MessageBoxFinal = "";
            var confirmBusNumber2 = confirmBusNumber.IndexOf("Bus ID: ");
            var confirmBusNumberPre = confirmBusNumber.Substring(confirmBusNumber2 + 8, 4);
            var confirmBusNumberFinal = Regex.Replace(confirmBusNumberPre, "[^0-9.]", "");
            //start
            Process p2 = new Process();
            p2.StartInfo.UseShellExecute = false;
            p2.StartInfo.RedirectStandardOutput = true;
            p2.StartInfo.CreateNoWindow = true;
            p2.StartInfo.FileName = @"ecs87NextGearOD.exe";
            p2.StartInfo.Arguments = "b " + confirmBusNumberFinal;
            p2.Start();
            Thread.Sleep(500);
            while ((standard_output = p2.StandardOutput.ReadLine()) != null)
            {
                MessageBoxFinal += standard_output + "\n";
            }
            MessageBox.Show(MessageBoxFinal);
            p2.Close();
        }
        private void buttonThermals_Click(object sender, RoutedEventArgs e)
        {
            if (min_FanSpeed.Text == "" || target_FanSpeed.Text == "" | target_Temperature.Text == "")
            {

            }
            else
            {
                //get bus number
                string confirmBusNumber = deviceComboBox.Text;
                var confirmBusNumber2 = confirmBusNumber.IndexOf("Bus ID: ");
                var confirmBusNumberPre = confirmBusNumber.Substring(confirmBusNumber2 + 8, 4);
                var confirmBusNumberFinal = Regex.Replace(confirmBusNumberPre, "[^0-9.]", "");
                //start
                Process p2 = new Process();
                p2.StartInfo.UseShellExecute = false;
                p2.StartInfo.RedirectStandardOutput = true;
                p2.StartInfo.CreateNoWindow = true;
                p2.StartInfo.FileName = @"ecs87NextGearOD.exe";
                p2.StartInfo.Arguments = "f 1 " + min_FanSpeed.Text + " " + confirmBusNumberFinal;
                p2.Start();
                Thread.Sleep(500);
                p2.Close();
                Process p1 = new Process();
                p1.StartInfo.UseShellExecute = false;
                p1.StartInfo.CreateNoWindow = true;
                p1.StartInfo.RedirectStandardOutput = true;
                p1.StartInfo.FileName = @"ecs87NextGearOD.exe";
                p1.StartInfo.Arguments = "f 2 " + target_FanSpeed.Text + " " + confirmBusNumberFinal;
                p1.Start();
                Thread.Sleep(500);
                p2.Close();
                Process p3 = new Process();
                p3.StartInfo.UseShellExecute = false;
                p3.StartInfo.CreateNoWindow = true;
                p3.StartInfo.RedirectStandardOutput = true;
                p3.StartInfo.FileName = @"ecs87NextGearOD.exe";
                p3.StartInfo.Arguments = "f 3 " + target_Temperature.Text + " " + confirmBusNumberFinal;
                p3.Start();
                Thread.Sleep(500);
                p3.Close();
            }
        }
        private void buttonReset_Click(object sender, System.EventArgs e)
        {
            try
            {
                //get bus number
                string confirmBusNumber = deviceComboBox.Text;
                var confirmBusNumber2 = confirmBusNumber.IndexOf("Bus ID: ");
                var confirmBusNumberPre = confirmBusNumber.Substring(confirmBusNumber2 + 8, 4);
                var confirmBusNumberFinal = Regex.Replace(confirmBusNumberPre, "[^0-9.]", "");
                //switch
                dropdown_reset.Items.Add("Clock Speeds");
                dropdown_reset.Items.Add("Voltages");
                dropdown_reset.Items.Add("All");
                if (dropdown_reset.Text == "All")
                {
                    //start
                    Process p1 = new Process();
                    p1.StartInfo.UseShellExecute = false;
                    p1.StartInfo.RedirectStandardOutput = true;
                    p1.StartInfo.CreateNoWindow = true;
                    p1.StartInfo.FileName = @"ecs87NextGearOD.exe";
                    //Reset everything
                    p1.StartInfo.Arguments = "rmc " + confirmBusNumberFinal;
                    p1.Start();
                    Thread.Sleep(500); //give it time to...breathe
                    p1.Close();
                }
                else if (dropdown_reset.Text == "Voltages")
                {
                    //save old clock speeds
                    var PPlayCC1_backup = PPlayCC1.Text;
                    var PPlayCC2_backup = PPlayCC2.Text;
                    var PPlayCC3_backup = PPlayCC3.Text;
                    var PPlayCC4_backup = PPlayCC4.Text;
                    var PPlayCC5_backup = PPlayCC5.Text;
                    var PPlayCC6_backup = PPlayCC6.Text;
                    var PPlayCC7_backup = PPlayCC7.Text;
                    var PPlayCC8_backup = PPlayCC8.Text;
                    var PPlayMC1_backup = PPlayMC1.Text;
                    var PPlayMC2_backup = PPlayMC2.Text;
                    var PPlayMC3_backup = PPlayMC3.Text;
                    var PPlayMC4_backup = PPlayMC4.Text;
                    var PPlayMC5_backup = PPlayMC5.Text;
                    var PPlayMC6_backup = PPlayMC6.Text;
                    var PPlayMC7_backup = PPlayMC7.Text;
                    var PPlayMC8_backup = PPlayMC8.Text;
                    //start
                    Process p1 = new Process();
                    p1.StartInfo.UseShellExecute = false;
                    p1.StartInfo.RedirectStandardOutput = true;
                    p1.StartInfo.CreateNoWindow = true;
                    p1.StartInfo.FileName = @"ecs87NextGearOD.exe";
                    //Reset everything
                    p1.StartInfo.Arguments = "rmc " + confirmBusNumberFinal;
                    p1.Start();
                    Thread.Sleep(500); //give it time to...breathe
                    p1.Close();
                    //start
                    Process p2 = new Process();
                    p2.StartInfo.UseShellExecute = false;
                    p2.StartInfo.RedirectStandardOutput = true;
                    p2.StartInfo.CreateNoWindow = true;
                    p2.StartInfo.FileName = @"ecs87NextGearOD.exe";
                    //Set backuped clocks
                    p2.StartInfo.Arguments = "g setclocksandvoltages " + PPlayCC1_backup + " " + PPlayCC2_backup + " " + PPlayCC3_backup + " " + PPlayCC4_backup + " " + PPlayCC5_backup + " " + PPlayCC6_backup + " " + PPlayCC7_backup + " " + PPlayCC8_backup + " " +
                        PPlayMC1_backup + " " + PPlayMC2_backup + " " + PPlayMC3_backup + " " + PPlayMC4_backup + " " + PPlayMC5_backup + " " + PPlayMC6_backup + " " + PPlayMC7_backup + " " + PPlayMC8_backup + " " +
                        PPlayCV1.Text + " " + PPlayCV2.Text + " " + PPlayCV3.Text + " " + PPlayCV4.Text + " " + PPlayCV5.Text + " " + PPlayCV6.Text + " " + PPlayCV7.Text + " " + PPlayCV8.Text + " " +
                        PPlayMV1.Text + " " + PPlayMV2.Text + " " + PPlayMV3.Text + " " + PPlayMV4.Text + " " + PPlayMV5.Text + " " + PPlayMV6.Text + " " + PPlayMV7.Text + " " + PPlayMV8.Text + " " +
                        confirmBusNumberFinal;
                    p2.Start();
                    Thread.Sleep(500); //give it time to...breathe
                    p2.Close();
                }
                else if (dropdown_reset.Text == "Clocks")
                {
                    //save old clock speeds
                    var PPlayCV1_backup = PPlayCV1.Text;
                    var PPlayCV2_backup = PPlayCV2.Text;
                    var PPlayCV3_backup = PPlayCV3.Text;
                    var PPlayCV4_backup = PPlayCV4.Text;
                    var PPlayCV5_backup = PPlayCV5.Text;
                    var PPlayCV6_backup = PPlayCV6.Text;
                    var PPlayCV7_backup = PPlayCV7.Text;
                    var PPlayCV8_backup = PPlayCV8.Text;
                    var PPlayMV1_backup = PPlayMV1.Text;
                    var PPlayMV2_backup = PPlayMV2.Text;
                    var PPlayMV3_backup = PPlayMV3.Text;
                    var PPlayMV4_backup = PPlayMV4.Text;
                    var PPlayMV5_backup = PPlayMV5.Text;
                    var PPlayMV6_backup = PPlayMV6.Text;
                    var PPlayMV7_backup = PPlayMV7.Text;
                    var PPlayMV8_backup = PPlayMV8.Text;
                    //start
                    Process p1 = new Process();
                    p1.StartInfo.UseShellExecute = false;
                    p1.StartInfo.RedirectStandardOutput = true;
                    p1.StartInfo.CreateNoWindow = true;
                    p1.StartInfo.FileName = @"ecs87NextGearOD.exe";
                    //Reset everything
                    p1.StartInfo.Arguments = "rmc " + confirmBusNumberFinal;
                    p1.Start();
                    Thread.Sleep(500); //give it time to...breathe
                    p1.Close();
                    //start
                    Process p2 = new Process();
                    p2.StartInfo.UseShellExecute = false;
                    p2.StartInfo.RedirectStandardOutput = true;
                    p2.StartInfo.CreateNoWindow = true;
                    p2.StartInfo.FileName = @"ecs87NextGearOD.exe";
                    //Set backuped clocks
                    p2.StartInfo.Arguments = "g setclocksandvoltages " + PPlayCC1.Text + " " + PPlayCC2.Text + " " + PPlayCC3.Text + " " + PPlayCC4.Text + " " + PPlayCC5.Text + " " + PPlayCC6.Text + " " + PPlayCC7.Text + " " + PPlayCC8.Text + " " +
                        PPlayMC1.Text + " " + PPlayMC2.Text + " " + PPlayMC3.Text + " " + PPlayMC4.Text + " " + PPlayMC5.Text + " " + PPlayMC6.Text + " " + PPlayMC7.Text + " " + PPlayMC8.Text + " " +
                        PPlayCV1_backup + " " + PPlayCV2_backup + " " + PPlayCV3_backup + " " + PPlayCV4_backup + " " + PPlayCV5_backup + " " + PPlayCV6_backup + " " + PPlayCV7_backup + " " + PPlayCV8_backup + " " +
                        PPlayMV1_backup + " " + PPlayMV2_backup + " " + PPlayMV3_backup + " " + PPlayMV4_backup + " " + PPlayMV5_backup + " " + PPlayMV6_backup + " " + PPlayMV7_backup + " " + PPlayMV8_backup + " " +
                        confirmBusNumberFinal;
                    p2.Start();
                    Thread.Sleep(500); //give it time to...breathe
                    p2.Close();
                }
            }
            catch { }
        }
        private void buttonClkAndVolts_Click(object sender, RoutedEventArgs e)
        {
            if (min_FanSpeed.Text == "" || target_FanSpeed.Text == "" | target_Temperature.Text == "")
            {

            }
            else
            {
                //get bus number
                string confirmBusNumber = deviceComboBox.Text;
                var confirmBusNumber2 = confirmBusNumber.IndexOf("Bus ID: ");
                var confirmBusNumberPre = confirmBusNumber.Substring(confirmBusNumber2 + 8, 4);
                var confirmBusNumberFinal = Regex.Replace(confirmBusNumberPre, "[^0-9.]", "");
                //start
                Process p2 = new Process();
                p2.StartInfo.UseShellExecute = false;
                p2.StartInfo.RedirectStandardOutput = true;
                p2.StartInfo.CreateNoWindow = true;
                p2.StartInfo.FileName = @"ecs87NextGearOD.exe";
                p2.StartInfo.Arguments = "g setclocksandvoltages " + PPlayCC1.Text + " " + PPlayCC2.Text + " " + PPlayCC3.Text + " " + PPlayCC4.Text + " " + PPlayCC5.Text + " " + PPlayCC6.Text + " " + PPlayCC7.Text + " " + PPlayCC8.Text + " " +
                    PPlayMC1.Text + " " + PPlayMC2.Text + " " + PPlayMC3.Text + " " + PPlayMC4.Text + " " + PPlayMC5.Text + " " + PPlayMC6.Text + " " + PPlayMC7.Text + " " + PPlayMC8.Text + " " +
                    PPlayCV1.Text + " " + PPlayCV2.Text + " " + PPlayCV3.Text + " " + PPlayCV4.Text + " " + PPlayCV5.Text + " " + PPlayCV6.Text + " " + PPlayCV7.Text + " " + PPlayCV8.Text + " " +
                    PPlayMV1.Text + " " + PPlayMV2.Text + " " + PPlayMV3.Text + " " + PPlayMV4.Text + " " + PPlayMV5.Text + " " + PPlayMV6.Text + " " + PPlayMV7.Text + " " + PPlayMV8.Text + " " +
                    confirmBusNumberFinal;
                p2.Start();
                Thread.Sleep(500); //give it time to...breathe
                p2.Close();
            }
        }
        private void buttonClkAndVolts_Save_Click(object sender, System.EventArgs e)
        {
            try
            {
                Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
                dlg.FileName = "NextGenODProfile"; // Default file name
                dlg.DefaultExt = ".ngOD"; // Default file extension
                Nullable<bool> result = dlg.ShowDialog();
                //get bus number
                string confirmBusNumber = deviceComboBox.Text;
                var confirmBusNumber2 = confirmBusNumber.IndexOf("Bus ID: ");
                var confirmBusNumberPre = confirmBusNumber.Substring(confirmBusNumber2 + 8, 4);
                var confirmBusNumberFinal = Regex.Replace(confirmBusNumberPre, "[^0-9.]", "");
                //get clocks and voltages
                string clockandvoltages = "g setclocksandvoltages " + PPlayCC1.Text + " " + PPlayCC2.Text + " " + PPlayCC3.Text + " " + PPlayCC4.Text + " " + PPlayCC5.Text + " " + PPlayCC6.Text + " " + PPlayCC7.Text + " " + PPlayCC8.Text + " " +
                        PPlayMC1.Text + " " + PPlayMC2.Text + " " + PPlayMC3.Text + " " + PPlayMC4.Text + " " + PPlayMC5.Text + " " + PPlayMC6.Text + " " + PPlayMC7.Text + " " + PPlayMC8.Text + " " +
                        PPlayCV1.Text + " " + PPlayCV2.Text + " " + PPlayCV3.Text + " " + PPlayCV4.Text + " " + PPlayCV5.Text + " " + PPlayCV6.Text + " " + PPlayCV7.Text + " " + PPlayCV8.Text + " " +
                        PPlayMV1.Text + " " + PPlayMV2.Text + " " + PPlayMV3.Text + " " + PPlayMV4.Text + " " + PPlayMV5.Text + " " + PPlayMV6.Text + " " + PPlayMV7.Text + " " + PPlayMV8.Text + "\n";
                string fanControl1 = "f 2 " + target_FanSpeed.Text + "\n";
                string fanControl2 = "f 3 " + target_Temperature.Text + "\n";
                string fanControl3 = "f 1 " + min_FanSpeed.Text;
                // Process save file dialog box results
                if (result == true)
                {
                    File.WriteAllText(dlg.FileName, clockandvoltages);
                    File.AppendAllText(dlg.FileName, fanControl1);
                    File.AppendAllText(dlg.FileName, fanControl2);
                    File.AppendAllText(dlg.FileName, fanControl3);
                }
            }
            catch { }
        }
        private void buttonClkAndVolts_Load_Click(object sender, System.EventArgs e)
        {
            try
            {
                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.Title = "Open ngOD File"; // Default file name
                dlg.Filter = "Next Gen OD files only|*.ngOD"; // Filter files by extension
                Nullable<bool> FOresult = dlg.ShowDialog();
                if (FOresult == true)
                {
                    System.Collections.Generic.IEnumerable<String> lines = File.ReadLines(dlg.FileName);
                    foreach (string lineItem in lines)
                    {
                        //get bus number
                        string confirmBusNumber = deviceComboBox.Text;
                        var confirmBusNumber2 = confirmBusNumber.IndexOf("Bus ID: ");
                        var confirmBusNumberPre = confirmBusNumber.Substring(confirmBusNumber2 + 8, 4);
                        var confirmBusNumberFinal = Regex.Replace(confirmBusNumberPre, "[^0-9.]", "");
                        //start
                        Process p2 = new Process();
                        p2.StartInfo.UseShellExecute = false;
                        p2.StartInfo.RedirectStandardOutput = true;
                        p2.StartInfo.CreateNoWindow = true;
                        p2.StartInfo.FileName = @"ecs87NextGearOD.exe";
                        p2.StartInfo.Arguments = lineItem + " " + confirmBusNumberFinal;
                        p2.Start();
                        Thread.Sleep(100); //give it time to...breathe
                        p2.Close();
                    }
                    //refresh the stuff
                    string CurrentItem = (string)deviceComboBox.SelectedItem;
                    button_refreshDeviceCombobox_Click(this, null);
                    deviceComboBox.SelectedIndex = deviceComboBox.Items.IndexOf(CurrentItem);
                    var result = string.Join(" ", list.ToArray());
                    var stringList2 = Regex.Split(result, "Bus ID:").ToList();
                    foreach (var stringlistitem in stringList2)
                    {
                        var stringlistitemfinal = " Bus ID:" + stringlistitem;
                        try
                        {
                            if (stringlistitemfinal.Contains(CurrentItem) && stringlistitemfinal.Contains("Assigned Target Temperature :"))
                            {
                                fansAndTemps(stringlistitem);
                            }
                            if (stringlistitemfinal.Contains(CurrentItem))
                            {
                                loopClocksAndVoltages(stringlistitem, 0);
                                loopClocksAndVoltages(stringlistitem, 1);
                                loopClocksAndVoltages(stringlistitem, 2);
                                loopClocksAndVoltages(stringlistitem, 3);
                                loopClocksAndVoltages(stringlistitem, 4);
                                loopClocksAndVoltages(stringlistitem, 5);
                                loopClocksAndVoltages(stringlistitem, 6);
                                loopClocksAndVoltages(stringlistitem, 7);
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
        }
        private void button_VoltageOffset_Click(object sender, RoutedEventArgs e)
        {
            if (cmVoffset.Text == "")
            {

            }
            else
            {
                try
                {
                    //Core Voltage Offsets
                    double[] CVmod = new double[8];
                    int[] cVoffsetFinal = new int[8];
                    CVmod[0] = Convert.ToInt32(PPlayCV1.Text);
                    CVmod[1] = Convert.ToInt32(PPlayCV2.Text);
                    CVmod[2] = Convert.ToInt32(PPlayCV3.Text);
                    CVmod[3] = Convert.ToInt32(PPlayCV4.Text);
                    CVmod[4] = Convert.ToInt32(PPlayCV5.Text);
                    CVmod[5] = Convert.ToInt32(PPlayCV6.Text);
                    CVmod[6] = Convert.ToInt32(PPlayCV7.Text);
                    CVmod[7] = Convert.ToInt32(PPlayCV8.Text);
                    int cVOffsetInt = Convert.ToInt32(cmVoffset.Text);
                    for (int i = 0; i < CVmod.Length; i++)
                    {
                        if (CVmod[i] == 0) { continue; }
                        CVmod[i] = CVmod[i] + (cVOffsetInt * 6.25);
                        cVoffsetFinal[i] = (int)Math.Round(CVmod[i], 0);
                    }
                    PPlayCV1.Text = Convert.ToString(cVoffsetFinal[0]);
                    PPlayCV2.Text = Convert.ToString(cVoffsetFinal[1]);
                    PPlayCV3.Text = Convert.ToString(cVoffsetFinal[2]);
                    PPlayCV4.Text = Convert.ToString(cVoffsetFinal[3]);
                    PPlayCV5.Text = Convert.ToString(cVoffsetFinal[4]);
                    PPlayCV6.Text = Convert.ToString(cVoffsetFinal[5]);
                    PPlayCV7.Text = Convert.ToString(cVoffsetFinal[6]);
                    PPlayCV8.Text = Convert.ToString(cVoffsetFinal[7]);
                    //Mem Voltage Offsets
                    double[] MVmod = new double[8];
                    int[] mVoffsetFinal = new int[8];
                    MVmod[0] = Convert.ToInt32(PPlayMV1.Text);
                    MVmod[1] = Convert.ToInt32(PPlayMV2.Text);
                    MVmod[2] = Convert.ToInt32(PPlayMV3.Text);
                    MVmod[3] = Convert.ToInt32(PPlayMV4.Text);
                    MVmod[4] = Convert.ToInt32(PPlayMV5.Text);
                    MVmod[5] = Convert.ToInt32(PPlayMV6.Text);
                    MVmod[6] = Convert.ToInt32(PPlayMV7.Text);
                    MVmod[7] = Convert.ToInt32(PPlayMV8.Text);
                    int mVOffsetInt = Convert.ToInt32(cmVoffset.Text);
                    for (int i = 0; i < MVmod.Length; i++)
                    {
                        if (MVmod[i] == 0) { continue; }
                        MVmod[i] = MVmod[i] + (mVOffsetInt * 6.25);
                        mVoffsetFinal[i] = (int)Math.Round(MVmod[i], 0);
                    }
                    PPlayMV1.Text = Convert.ToString(mVoffsetFinal[0]);
                    PPlayMV2.Text = Convert.ToString(mVoffsetFinal[1]);
                    PPlayMV3.Text = Convert.ToString(mVoffsetFinal[2]);
                    PPlayMV4.Text = Convert.ToString(mVoffsetFinal[3]);
                    PPlayMV5.Text = Convert.ToString(mVoffsetFinal[4]);
                    PPlayMV6.Text = Convert.ToString(mVoffsetFinal[5]);
                    PPlayMV7.Text = Convert.ToString(mVoffsetFinal[6]);
                    PPlayMV8.Text = Convert.ToString(mVoffsetFinal[7]);
                    //buttonClkAndVolts_Click(this, null);
                }
                catch { }
            }
        }
    }
}
