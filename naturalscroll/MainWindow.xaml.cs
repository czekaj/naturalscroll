using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.Collections;
using System.Security;

namespace naturalscroll
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        Hashtable pointingDevices = new Hashtable();

        private void deviceList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            String registrySubKey = "SYSTEM\\CurrentControlSet\\Enum\\" + ((DictionaryEntry)deviceList.SelectedItem).Key + "\\Device Parameters";
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(registrySubKey))
                {
                    if (key != null)
                    {
                        Object o = key.GetValue("FlipFlopWheel");
                        if (o != null)
                        {
                            naturalScrollCheckBox.IsChecked = o.ToString() == "1" ? true : false;
                            naturalScrollCheckBox.IsEnabled = true;
                            saveButton.IsEnabled = true;
                        }
                        else
                        {
                            saveButton.IsEnabled = false;
                            naturalScrollCheckBox.IsEnabled = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void deviceList_Loaded(object sender, RoutedEventArgs e)
        {
            using (PowerShell PowerShellInstance = PowerShell.Create())
            {
                PowerShellInstance.AddScript("Get-WmiObject win32_PointingDevice");

                Collection<PSObject> PSOutput = PowerShellInstance.Invoke();
                foreach (PSObject outputItem in PSOutput)
                {
                    // if null object was dumped to the pipeline during the script then a null
                    // object may be present here. check for null to prevent potential NRE.
                    if (outputItem != null)
                    {
                        var element = outputItem.BaseObject as ManagementObject;
                        pointingDevices.Add(element.GetPropertyValue("DeviceID"), element.GetPropertyValue("Caption"));
                    }
                }
            }
            deviceList.ItemsSource = pointingDevices;
            deviceList.DisplayMemberPath = "Value";
        }

        private void saveButton_Click(object sender, RoutedEventArgs e)
        {
            String registrySubKey = "SYSTEM\\CurrentControlSet\\Enum\\" + ((DictionaryEntry)deviceList.SelectedItem).Key + "\\Device Parameters";
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(registrySubKey, true))
                {
                    if (key != null)
                    {
                        Object o = key.GetValue("FlipFlopWheel");
                        if (o != null)
                        {
                            key.SetValue("FlipFlopWheel", naturalScrollCheckBox.IsChecked == true ? 1 : 0);
                            naturalScrollCheckBox.IsEnabled = false;
                            saveButton.IsEnabled = false;
                            MessageBox.Show("Your changes have been saved. Please reboot your machine to apply.");
                        }
                        else
                        {
                            naturalScrollCheckBox.IsEnabled = false;
                            saveButton.IsEnabled = false;
                        }
                    }
                }
            }
            catch (SecurityException sex) 
            {
                MessageBox.Show("Unfortunately, the application doesn't have access to system registry and cannot operate properly.\n(" + sex.Message + ")\n\nPlease run the application again as Administrator.");
            }

        }
    }
}
