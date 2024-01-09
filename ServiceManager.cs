using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Windows.Forms;

namespace Service_Manager
{
    public partial class ServiceManager
    {
        private string FilePath;

        public ServiceManager()
        {
            FilePath = Environment.CurrentDirectory + @"\ServiceManager.ini";
            if (File.Exists(FilePath))
                GetServiceGroups();
            else
                LoadDefaultServiceGroup();
        }

        public List<string> GetServiceGroups()
        {
            List<string> serviceGroupNames = new List<string>();
            try
            {
                StreamReader ReadFile_ServiceGroup = new StreamReader(FilePath, Encoding.GetEncoding(936));
                string line;
                while ((line = ReadFile_ServiceGroup.ReadLine()) != null)
                    if (line.StartsWith("#"))
                        serviceGroupNames.Add(line.TrimStart('#'));
                ReadFile_ServiceGroup.Close();
            }
            catch (IOException)
            {
                Error_DBFile();
                return null;
            }
            return serviceGroupNames;
        }
        public List<string> GetServiceGroup(string groupName)
        {
            List<string> serviceNames = new List<string>();

            try
            {
                StreamReader ReadFile_ServiceGroup = new StreamReader(FilePath, Encoding.GetEncoding(936));
                string line;
                while ((line = ReadFile_ServiceGroup.ReadLine()) != null)
                    if (line.StartsWith("#") && line.TrimStart('#') == groupName)
                    {
                        int recordCnt = Convert.ToInt32(ReadFile_ServiceGroup.ReadLine());
                        for (int i = 1; i <= recordCnt; i++)
                        {
                            line = ReadFile_ServiceGroup.ReadLine();
                            if (line == null) throw new IOException();
                            else serviceNames.Add(line);
                        }
                        break;
                    }
                ReadFile_ServiceGroup.Close();
            }
            catch (IOException)
            {
                Error_DBFile();
                return null;
            }

            return serviceNames;
        }
        public bool AddServiceGroup(string groupName, List<string> serviceNames)
        {
            try
            {
                StreamWriter WriteFile_ServiceGroup = new StreamWriter(FilePath, true, Encoding.GetEncoding(936));
                WriteFile_ServiceGroup.Write("#" + groupName + "\r\n");
                WriteFile_ServiceGroup.Write(serviceNames.Count.ToString() + "\r\n");
                foreach (string serviceName in serviceNames)
                    WriteFile_ServiceGroup.Write(serviceName + "\r\n");
                WriteFile_ServiceGroup.Write("\r\n");
                WriteFile_ServiceGroup.Close();
            }
            catch (IOException)
            {
                Error_DBFile();
                return false;
            }
            return true;
        }
        public bool DeleteServiceGroup(string groupName)
        {
            try
            {
                StreamReader ReadFile_ServiceGroup = new StreamReader(FilePath, Encoding.GetEncoding(936));
                string line;
                int ptr = -1;
                while ((line = ReadFile_ServiceGroup.ReadLine()) != null)
                {
                    ptr++;
                    if (line.StartsWith("#") && line.TrimStart('#') == groupName)
                    {
                        int recordCnt = Convert.ToInt32(ReadFile_ServiceGroup.ReadLine());
                        recordCnt += 3;

                        ReadFile_ServiceGroup.Close();

                        List<string> lines = new List<string>(File.ReadAllLines(FilePath, Encoding.GetEncoding(936)));
                        for (int i = 0; i < recordCnt; i++)
                            lines.RemoveAt(ptr);
                        File.WriteAllLines(FilePath, lines.ToArray(), Encoding.GetEncoding(936));
                        return true;
                    }
                }
                ReadFile_ServiceGroup.Close();
            }
            catch (IOException)
            {
                Error_DBFile();
                return false;
            }
            return false;
        }
        public void EditDB()
        {
            ProcessStartInfo info = new ProcessStartInfo("Notepad.exe", FilePath);
            Process.Start(info);
        }
        public void ResetDB()
        {
            LoadDefaultServiceGroup();
        }

        private void LoadDefaultServiceGroup()
        {
            StreamWriter WriteFile_ServiceGroup = new StreamWriter(FilePath, false, Encoding.GetEncoding(936));

            string defaultServiceGroup = "#SQL Server\r\n5\r\nMSSQL$KAHSOLT\r\nSQLBrowser\r\nSQLTELEMETRY$KAHSOLT\r\nSQLWriter\r\nSQLAgent$KAHSOLT\r\n\r\n" +
                "#NVIDIA\r\n5\r\nnvsvc\r\nGfExperienceService\r\nNvNetworkService\r\nNvStreamNetworkSvc\r\nNvStreamSvc\r\n\r\n" +
                "#Hyper-V\r\n8\r\nvmickvpexchange\r\nvmicguestinterface\r\nvmicshutdown\r\nvmicheartbeat\r\nvmicvmsession\r\nvmictimesync\r\nvmicvss\r\nvmicrdv\r\n\r\n" +
                "#VMware\r\n4\r\nVMAuthdService\r\nVMnetDHCP\r\nVMware NAT Service\r\nVMUSBArbService\r\n\r\n" +
                "#Windows Defender\r\n3\r\nSense\r\nWdNisSvc\r\nWinDefend\r\n\r\n" +
                "#Microsoft Office\r\n2\r\nClickToRunSvc\r\nose64\r\n\r\n" +
                "#Rosetta Stone\r\n1\r\nRosettaStoneDaemon\r\n\r\n" +
                "#Thunder\r\n1\r\nXLServicePlatform\r\n\r\n" +
                "#Tencent\r\n2\r\nQPCore\r\nQQMusicService\r\n\r\n";

            WriteFile_ServiceGroup.Write(defaultServiceGroup);
            WriteFile_ServiceGroup.Close();
        }
        private void Error_DBFile()
        {
            if (MessageBox.Show("Fatal Error: Помилка читання або запису файлу даних!\nЧи хочете ви відновити інформацію про службову групу за замовчуванням?", "Системна помилка!", MessageBoxButtons.OKCancel) == DialogResult.OK)
                LoadDefaultServiceGroup();
        }

        //Service & Device
        public int GetServices(ListView serviceInfo)
        {
            ServiceController[] serviceControllers = ServiceController.GetServices();

            serviceInfo.BeginUpdate();
            serviceInfo.Items.Clear();
            foreach (ServiceController service in serviceControllers)
            {
                ListViewItem item = new ListViewItem();
                item.Text = service.DisplayName;
                item.SubItems.Add(service.ServiceName);
                item.SubItems.Add(getStatus(service));
                item.SubItems.Add(getStartupType(service));
                item.SubItems.Add(getCompany(service));
                item.SubItems.Add(getDescription(service));
                item.SubItems.Add(getImageCommands(service.ServiceName));
                serviceInfo.Items.Add(item);
                service.Close();
            }
            serviceInfo.EndUpdate();
            return serviceInfo.Items.Count;
        }
        public void GetDependingInfo(string serviceName, ref List<string> servicesDependedOn, ref List<string> dependentServices)
        {
            ServiceController[] services = ServiceController.GetServices();
            ServiceController thisService = null;

            foreach (ServiceController service in services)
                if (service.ServiceName == serviceName)
                {
                    thisService = service;
                    break;
                }
                else
                    service.Close();

            if (thisService == null)
            {
                MessageBox.Show("Помилка: немає такої служби!", "Системна помилка!");
                return;
            }

            services = thisService.ServicesDependedOn;
            foreach (ServiceController service in services)
                servicesDependedOn.Add(service.ServiceName);
            services = thisService.DependentServices;
            foreach (ServiceController service in services)
                dependentServices.Add(service.ServiceName);
        }
        public bool GetStatus(string serviceName)
        {
            ServiceController[] services = ServiceController.GetServices();
            ServiceController thisService = null;

            foreach (ServiceController service in services)
                if (service.ServiceName == serviceName)
                {
                    thisService = service;
                    break;
                }
                else
                    service.Close();

            if (thisService == null)
            {
                MessageBox.Show("Помилка: Немає такої служби!", "Системна помилка!");
                return false;
            }
            else if (thisService.Status != ServiceControllerStatus.Running)
                return false;
            else
                return true;
        }
        public bool StartService(string serviceName)
        {
            try
            {
                ServiceController serviceController = new ServiceController(serviceName);
                serviceController.Start();
                serviceController.WaitForStatus(ServiceControllerStatus.Running);
                serviceController.Close();
            }
            catch
            {
                return false;
            }
            return true;
        }
        public bool RestartService(string serviceName)
        {
            try
            {
                ServiceController serviceController = new ServiceController(serviceName);
                serviceController.Stop();
                serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                serviceController.Start();
                serviceController.WaitForStatus(ServiceControllerStatus.Running);
                serviceController.Close();
            }
            catch
            {
                return false;
            }
            return true;
        }
        public bool StopService(string serviceName)
        {
            try
            {
                ServiceController serviceController = new ServiceController(serviceName);
                serviceController.Stop();
                serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                serviceController.Close();
            }
            catch
            {
                return false;
            }
            return true;
        }
        public bool ChangeStartupType(string serviceName, int startupType)
        {
            try
            {
                switch (startupType)
                {
                    case 2:
                        Registry.LocalMachine.CreateSubKey(@"SYSTEM\CurrentControlSet\Services\" + serviceName).SetValue("Start", 2, RegistryValueKind.DWord);
                        break;
                    case 3:
                        Registry.LocalMachine.CreateSubKey(@"SYSTEM\CurrentControlSet\Services\" + serviceName).SetValue("Start", 3, RegistryValueKind.DWord);
                        break;
                    case 4:
                        Registry.LocalMachine.CreateSubKey(@"SYSTEM\CurrentControlSet\Services\" + serviceName).SetValue("Start", 4, RegistryValueKind.DWord);
                        break;
                    default:
                        MessageBox.Show("Помилка: sturtupType!", "Системна помилка!");
                        break;
                }

            }
            catch (System.Security.SecurityException)
            {
                MessageBox.Show("Помилка: не вдалося змінити!", "Підтвердження зміни типу запуску");
                return false;
            }
            return true;
        }
        public void ShowImageAttribute(string serviceName)
        {
            try
            {
                string fileName = getImagePath(serviceName);

                SHELLEXECUTEINFO info = new SHELLEXECUTEINFO();
                info.cbSize = Marshal.SizeOf(info);
                info.lpVerb = "properties";
                info.lpFile = fileName;
                info.nShow = 5;
                info.fMask = 12u;
                ShellExecuteEx(ref info);
            }
            catch
            {
                MessageBox.Show("Помилка: доступ відмовлено!", "Повідомлення про перегляд властивостей файлу.");
                return;
            }
        }
        public void ShowImageDirectory(string serviceName)
        {
            ProcessStartInfo fileInDirectory = new ProcessStartInfo("Explorer.exe");
            fileInDirectory.Arguments = "/e,/select," + getImagePath(serviceName);
            Process.Start(fileInDirectory);
        }

        private string getStatus(ServiceController service)
        {
            if (service.Status == ServiceControllerStatus.Running)
                return "Запущено";
            else if (service.Status == ServiceControllerStatus.Stopped)
                return "Завершено";
            else if (service.Status == ServiceControllerStatus.Paused)
                return "Зупинено";
            else if (service.Status == ServiceControllerStatus.PausePending)
                return "Призупинено";
            else if (service.Status == ServiceControllerStatus.ContinuePending)
                return "Триває відновлення";
            else if (service.Status == ServiceControllerStatus.StartPending)
                return "Запускається...";
            else if (service.Status == ServiceControllerStatus.StopPending)
                return "Зупиняється...";
            return "Невідомо";
        }
        private string getStartupType(ServiceController service)
        {
            try
            {
                string startupType = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services\" + service.ServiceName).GetValue("Start").ToString();
                switch (Convert.ToInt32(startupType))
                {
                    case 0:
                        return "Ядро";//Boot(by Kernel)
                    case 1:
                        return "Система";//System(by I/O Sub-System)
                    case 2:
                        return "Автоматично";
                    case 3:
                        return "Вручну";
                    case 4:
                        return "Вимкнено";
                    default:
                        return "Невідомо";
                }
            }
            catch
            {
                return "(Помилка доступу)";
            }
        }
        private string getCompany(ServiceController service)
        {
            string imageName = getImagePath(service.ServiceName);
            try
            {
                return FileVersionInfo.GetVersionInfo(imageName).CompanyName;
            }
            catch
            {
                return "";
            }
        }
        private string getDescription(ServiceController service)
        {
            try
            {
                string indirectString = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services\" + service.ServiceName).GetValue("Description").ToString();
                if (indirectString.StartsWith("@"))
                {
                    uint buffer = 1024;
                    char[] szOut = new char[buffer];
                    SHLoadIndirectString(indirectString.ToCharArray(), szOut, buffer, null);
                    return new string(szOut);
                }
                else
                    return indirectString;
            }
            catch
            {
                return "";
            }
        }
        private string getImageCommands(string serviceName)
        {
            try
            {
                return Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services\" + serviceName).GetValue("ImagePath").ToString();
            }
            catch
            {
                return "";
            }
        }
        private string getImagePath(string serviceName)
        {
            try
            {
                string imageName = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services\" + serviceName).GetValue("ImagePath").ToString();
                if (imageName.StartsWith("\""))
                    imageName = imageName.Substring(1, imageName.LastIndexOf('\"') - 1);
                else
                    imageName = imageName.Split('-')[0].Split('/')[0].Trim('"').TrimEnd(' ');
                if (!imageName.EndsWith(".exe"))
                    imageName += ".exe";
                return imageName;
            }
            catch
            {
                return "";
            }
        }


        #region Call DLLs - 支持文件属性
        [StructLayout(LayoutKind.Sequential)]
        public struct SHELLEXECUTEINFO
        {
            public int cbSize;
            public uint fMask;
            public IntPtr hwnd;
            [MarshalAs(UnmanagedType.LPStr)]
            public string lpVerb;
            [MarshalAs(UnmanagedType.LPStr)]
            public string lpFile;
            [MarshalAs(UnmanagedType.LPStr)]
            public string lpParameters;
            [MarshalAs(UnmanagedType.LPStr)]
            public string lpDirectory;
            public int nShow;
            public IntPtr hInstApp;
            public IntPtr lpIDList;
            [MarshalAs(UnmanagedType.LPStr)]
            public string lpClass;
            public IntPtr hkeyClass;
            public uint dwHotKey;
            public IntPtr hIcon;
            public IntPtr hProcess;
        }
        [DllImport("shell32.dll")]
        static extern bool ShellExecuteEx(ref SHELLEXECUTEINFO lpExecInfo);
        #endregion
        #region Call DLLs - 支持间接字符串转换
        [DllImport("Shlwapi.dll")]
        static extern int SHLoadIndirectString(char[] pszSource, char[] pszOutBuf, uint cchOutBuf, object ppvReserved);
        #endregion

    }
}
