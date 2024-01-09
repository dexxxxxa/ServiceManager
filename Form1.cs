using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Collections;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Service_Manager
{
    public partial class Form1 : Form
    {
        private const int ITEMS_COUNT = 7;
        private ToolStripMenuItem currentRefreshFrequency;
        private ToolStripMenuItem currentStartupType;
        private List<SortOrder> currentOrders;
        private bool filterEmpty;
        private static ServiceManager serviceManager = new ServiceManager();
        private static ListViewSorter listViewSorter = new ListViewSorter();


        public Form1()
        {
            InitializeComponent();
        }

        private void Form_Main_Load(object sender, EventArgs e)
        {
            currentRefreshFrequency = fastToolStripMenuItem;
            currentStartupType = toolStripMenuItem_Auto;
            currentOrders = new List<SortOrder>();
            for (int i = 1; i <= ITEMS_COUNT; i++)
                currentOrders.Add(SortOrder.Descending);
            this.listView_Service.ListViewItemSorter = listViewSorter;

            toolStripTextBox_Filter.Text = "Search";
            toolStripTextBox_Filter.Font = new Font(toolStripTextBox_Filter.Font, toolStripTextBox_Filter.Font.Style | FontStyle.Italic);
            filterEmpty = true;
            lightToolStripMenuItem.Checked = true;


            RefreshService();
            timer.Start();
            this.Focus();
        }

        protected void RefreshService()
        {
            if (filterEmpty)
            {
                serviceManager.GetServices(listView_Service);
                toolStripStatusLabel_Status.Text = "Доступно " + listView_Service.Items.Count.ToString() + " Елементів";
            }
            else
            {
                string filter = toolStripTextBox_Filter.Text;
                ListView filterList = new ListView();
                serviceManager.GetServices(filterList);

                listView_Service.BeginUpdate();
                listView_Service.Items.Clear();
                foreach (ListViewItem item in filterList.Items)
                    foreach (ListViewItem.ListViewSubItem subItem in item.SubItems)
                        if (subItem.Text.ToLowerInvariant().Contains(filter.ToLowerInvariant()))
                        {
                            listView_Service.Items.Add((ListViewItem)item.Clone());
                            break;
                        }
                listView_Service.EndUpdate();

                toolStripStatusLabel_Status.Text = "Доступно " + listView_Service.Items.Count.ToString() + " Елементів (Після фільтру“" + filter + "”)";
            }
        }

        private void listView_Service_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView_Service.FocusedItem == null)
                return;

            if (listView_Service.FocusedItem.SubItems[3].Text == "Завершено")
            {
                toolStripButtonStart.Enabled = false;
                toolStripButtonRestart.Enabled = false;
                toolStripButtonStop.Enabled = false;
                return;
            }

            switch (listView_Service.FocusedItem.SubItems[2].Text)
            {
                case "Запущено":
                    toolStripButtonStart.Enabled = false;
                    toolStripButtonRestart.Enabled = true;
                    toolStripButtonStop.Enabled = true;
                    break;
                case "Зупинено":
                    toolStripButtonStart.Enabled = true;
                    toolStripButtonRestart.Enabled = false;
                    toolStripButtonStop.Enabled = false;
                    break;
                default:
                    break;
            }
        }

        private void runToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process runDialog = new Process();
            runDialog.StartInfo.FileName = "C:\\Windows\\System32\\rundll32.exe";
            runDialog.StartInfo.Arguments = "shell32.dll,#61";
            runDialog.Start();
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string fileName;
            string seperator;

            fileName = "ServiceManager_";
            saveFileDialog_ListInfo.FileName = fileName + DateTime.Now.ToString("yyyy_MM_dd");
            saveFileDialog_ListInfo.ShowDialog();
            fileName = saveFileDialog_ListInfo.FileName;

            if (saveFileDialog_ListInfo.FileName.EndsWith(".csv"))
                seperator = ",";
            else
                seperator = "\t";

            StreamWriter streamWriter;
            if (File.Exists(fileName))
                streamWriter = new StreamWriter(fileName, false, Encoding.ASCII);
            else
                streamWriter = File.CreateText(fileName);
            for (int i = 0; i < listView_Service.Items.Count; i++)
            {
                for (int j = 0; j < listView_Service.Items[i].SubItems.Count; j++)
                    streamWriter.Write(listView_Service.Items[i].SubItems[j].Text + seperator);
                streamWriter.WriteLine();
            }
            streamWriter.Flush();
            streamWriter.Close();
        }

        private void screenshotToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IntPtr hWnd = this.Handle;
            IntPtr hscrdc = GetWindowDC(hWnd);
            Control control = FromHandle(hWnd);
            IntPtr hbitmap = CreateCompatibleBitmap(hscrdc, control.Width, control.Height);
            IntPtr hmemdc = CreateCompatibleDC(hscrdc);
            SelectObject(hmemdc, hbitmap);
            PrintWindow(hWnd, hmemdc, 0);
            Bitmap bmp = Image.FromHbitmap(hbitmap);
            DeleteDC(hscrdc);
            DeleteDC(hmemdc);

            string fileName;
            fileName = "ServiceManager_";
            saveFileDialog_SnapShot.FileName = fileName + DateTime.Now.ToString("yyyy_MM_dd");
            saveFileDialog_SnapShot.ShowDialog();
            fileName = saveFileDialog_SnapShot.FileName;
            if (fileName == "")
                return;


            if (saveFileDialog_SnapShot.FileName.EndsWith(".png"))
                bmp.Save(fileName, ImageFormat.Png);
            else if (saveFileDialog_SnapShot.FileName.EndsWith(".jpg"))
                bmp.Save(fileName, ImageFormat.Jpeg);
            else
                bmp.Save(fileName);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            notifyIcon.Visible = false;
            Application.Exit();
        }

        private void fastToolStripMenuItem_Click(object sender, EventArgs e)
        {
            currentRefreshFrequency.Checked = false;
            fastToolStripMenuItem.Checked = true;
            currentRefreshFrequency = fastToolStripMenuItem;
            ChangeTimerInterval();
        }

        private void slowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            currentRefreshFrequency.Checked = false;
            slowToolStripMenuItem.Checked = true;
            currentRefreshFrequency = slowToolStripMenuItem;
            ChangeTimerInterval();
        }

        private void stopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            currentRefreshFrequency.Checked = false;
            stopToolStripMenuItem.Checked = true;
            currentRefreshFrequency = stopToolStripMenuItem;
            ChangeTimerInterval();
        }

        protected void ChangeTimerInterval()
        {
            if (currentRefreshFrequency == stopToolStripMenuItem)
                timer.Stop();
            else
            {
                if (currentRefreshFrequency == slowToolStripMenuItem)
                    timer.Interval = 30000;
                else if (currentRefreshFrequency == fastToolStripMenuItem)
                    timer.Interval = 90000;
                else
                    MessageBox.Show("Помилка швидкості оновлення", "Помилка");
            }
        }

        private void toolStripButtonRefresh_Click(object sender, EventArgs e)
        {
            RefreshService();
        }

        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {
            serviceManager.EditDB();
        }

        private void setDefaultToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Ви впевнені, що хочете відновити конфігураційний файл за замовчуванням?", "Підтвердження відновлення...", MessageBoxButtons.OKCancel) == DialogResult.OK)
                serviceManager.ResetDB();
        }

        private void manualToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string userManual = "";
            userManual += "Інструкція з розширених функцій:\n\n";
            userManual += "1. Подвійний клік на будь-якому рядку в списку послуг дозволяє переглянути її спадковість.\n";
            userManual += "2. Клік на заголовку списку дозволяє сортувати його у зростаючому і спадному порядку.\n";
            userManual += "3. Список має контекстне меню за правою кнопкою миші, за замовчуванням відкриває іконку в системному сповіщенні.\n";
            userManual += "4. Формат файлу конфігурації групи послуг: назва групи з ведучим #, кількість послуг у групі та всі назви послуг у групі; розділовий рядок між групами.\n";
            userManual += "===============================================\n";
            userManual += "Відомі недоліки версії:\n\n";
            userManual += "1. Фільтр не завжди коректно працює, оскільки подія TextChange спрацьовує на кожний символ введеного тексту.\n";
            userManual += "2. Описи послуг часто порожні, оскільки C# не може безпосередньо зчитувати опорні рядки з ресурсів DLL-файлів.\n";
            userManual += "3. Перегляд властивостей файлу та визначення місця часом не працюють через використання функцій з бібліотеки C++, які, здається, не запускаються від імені адміністратора, що призводить до недостатніх прав доступу до Explorer.\n";
            userManual += "\n";

            MessageBox.Show(userManual, "Інструкція з використання...");
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Цей застосунок створив студент групи КІУКІ-21-10 - Даценко Денис",
                "Service Manager", MessageBoxButtons.OK);
        }

        public void listView_Service_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            if (currentOrders[e.Column] == SortOrder.Ascending)
                currentOrders[e.Column] = SortOrder.Descending;
            else
                currentOrders[e.Column] = SortOrder.Ascending;

            listViewSorter.SortColumn = e.Column;
            listViewSorter.SortOrder = currentOrders[e.Column];
            listView_Service.Sort();
        }

        private void toolStripTextBox_Filter_Enter(object sender, EventArgs e)
        {
            if (filterEmpty)
            {
                toolStripTextBox_Filter.Text = "";
                toolStripTextBox_Filter.Font = new Font(toolStripTextBox_Filter.Font, toolStripTextBox_Filter.Font.Style ^ FontStyle.Italic);
            }
        }

        private void toolStripTextBox_Filter_Leave(object sender, EventArgs e)
        {
            if (toolStripTextBox_Filter.Text == "")
            {
                filterEmpty = true;
                toolStripTextBox_Filter.Text = "Фільтр";
                toolStripTextBox_Filter.Font = new Font(toolStripTextBox_Filter.Font, toolStripTextBox_Filter.Font.Style | FontStyle.Italic);
            }
            else filterEmpty = false;
        }

        private void toolStripTextBox_Filter_TextChanged(object sender, EventArgs e)
        {
            if (!toolStripTextBox_Filter.Focused)
                return;
            else if (toolStripTextBox_Filter.Text == "" && filterEmpty)
                return;

           
            if (toolStripTextBox_Filter.Text == "")
                filterEmpty = true;
            else
                filterEmpty = false;

            RefreshService();
        }

        private void toolStripButtonStart_Click(object sender, EventArgs e)
        {
            if (listView_Service.SelectedItems.Count == 0)
                return;

            if (serviceManager.StartService(listView_Service.FocusedItem.SubItems[1].Text))
                listView_Service.FocusedItem.SubItems[2].Text = "Запущено";
            else
                MessageBox.Show("Не вдалося запустити.", "Помилка");

            listView_Service_SelectedIndexChanged(sender, e);
        }

        private void toolStripButtonRestart_Click(object sender, EventArgs e)
        {
            if (listView_Service.SelectedItems.Count == 0)
                return;
            if (serviceManager.RestartService(listView_Service.FocusedItem.SubItems[1].Text))
                listView_Service.FocusedItem.SubItems[2].Text = "Запущено";
            else
                MessageBox.Show("Не вдалося перезапустити.", "Помилка");

            listView_Service_SelectedIndexChanged(sender, e);
        }

        private void toolStripButtonStop_Click(object sender, EventArgs e)
        {
            if (listView_Service.SelectedItems.Count == 0)
                return;

            if (serviceManager.StopService(listView_Service.FocusedItem.SubItems[1].Text))
                listView_Service.FocusedItem.SubItems[2].Text = "Зупинено";
            else
                MessageBox.Show("Error: Не вдалося зупинити.", "Помилка");

            listView_Service_SelectedIndexChanged(sender, e);
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string serviceDisplayName = listView_Service.FocusedItem.SubItems[0].Text;
            Clipboard.SetDataObject(serviceDisplayName);
        }

        private void searchToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string serviceDisplayName = listView_Service.FocusedItem.SubItems[0].Text;
            string webpage = "https://www.google.com/search?q=Service%20" + serviceDisplayName;
            Process.Start(webpage);
        }

        private void propertiesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView_Service.SelectedItems.Count == 0)
                return;

            serviceManager.ShowImageAttribute(listView_Service.FocusedItem.SubItems[1].Text);
        }

        private void pathToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView_Service.SelectedItems.Count == 0)
                return;

            serviceManager.ShowImageDirectory(listView_Service.FocusedItem.SubItems[1].Text);
        }

        private void toolStripDropDownButton_Service_StartupType_DropDownOpening(object sender, EventArgs e)
        {
            if (listView_Service.SelectedItems.Count == 0)
                return;

            switch (listView_Service.FocusedItem.SubItems[3].Text)
            {
                case "Автоматично":
                    currentStartupType.Checked = false;
                    currentStartupType = toolStripMenuItem_Auto;
                    currentStartupType.Checked = true;
                    break;
                case "Вручну":
                    currentStartupType.Checked = false;
                    currentStartupType = toolStripMenuItem_Manual;
                    currentStartupType.Checked = true;
                    break;
                case "Зупинено":
                    currentStartupType.Checked = false;
                    currentStartupType = toolStripMenuItem_Disable;
                    currentStartupType.Checked = true;
                    break;
                default:
                    MessageBox.Show("Не вдалося змінити тип запуску.", "Помилка");
                    break;
            }
        }

        private void toolStripMenuItem_Manual_Click(object sender, EventArgs e)
        {
            if (serviceManager.ChangeStartupType(listView_Service.FocusedItem.SubItems[1].Text, 3))
                listView_Service.FocusedItem.SubItems[3].Text = "Вручну";
        }

        private void toolStripMenuItem_Auto_Click(object sender, EventArgs e)
        {
            if (serviceManager.ChangeStartupType(listView_Service.FocusedItem.SubItems[1].Text, 2))
                listView_Service.FocusedItem.SubItems[3].Text = "Автоматично";
        }

        private void toolStripMenuItem_Disabled_Click(object sender, EventArgs e)
        {
            if (serviceManager.ChangeStartupType(listView_Service.FocusedItem.SubItems[1].Text, 4))
                listView_Service.FocusedItem.SubItems[3].Text = "Зупинено";
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            RefreshService();
        }

        private void listView_Service_DoubleClick(object sender, EventArgs e)
        {
            if (listView_Service.FocusedItem == null)
                return;

            List<string> servicesDependedOn = new List<string>();
            List<string> dependentServices = new List<string>();
            serviceManager.GetDependingInfo(listView_Service.FocusedItem.SubItems[1].Text, ref servicesDependedOn, ref dependentServices);

            string dependingInfo = "=====Залежності служби " + listView_Service.FocusedItem.SubItems[1].Text + "=====\n\n";
            dependingInfo += "Ця служба залежить від (вищий рівень): " + servicesDependedOn.Count + " елементів\n";
            if (servicesDependedOn.Count == 0)
                dependingInfo += "(немає)\n";
            else
                foreach (string service in servicesDependedOn)
                    dependingInfo += service + "\n";
            dependingInfo += "\n";
            dependingInfo += "Залежать від цієї служби (нижчий рівень): " + dependentServices.Count + " елементів\n";
            if (dependentServices.Count == 0)
                dependingInfo += "(немає)\n";
            else
                foreach (string service in dependentServices)
                    dependingInfo += service + "\n";

            MessageBox.Show(dependingInfo, "Залежності служб...");
            return;

        }

        private void notifyIcon_DoubleClick(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.WindowState = FormWindowState.Normal;
                this.ShowInTaskbar = true;
                this.Activate();
            }
            else
            {
                this.WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
            }
        }

        private void lightThemeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (darkToolStripMenuItem.Checked == true)
            {
                this.BackColor = Color.White;
                this.ForeColor = Color.Black;

                toolStripTextBox_Filter.BackColor = Color.White;
                toolStripTextBox_Filter.ForeColor = Color.Black;

                listView_Service.BackColor = Color.White;
                listView_Service.ForeColor = Color.Black;
                listView_Service.Scrollable = true;
                listView_Service.HeaderStyle = ColumnHeaderStyle.Clickable;

                foreach (Control ctrl in this.Controls)
                {
                    ctrl.BackColor = Color.White;
                    ctrl.ForeColor = Color.Black;
                    if (ctrl is Button)
                    {
                        Button btn = ctrl as Button;
                        btn.FlatStyle = FlatStyle.Standard;
                        btn.BackColor = SystemColors.Control;
                        btn.ForeColor = Color.Black;
                    }
                }
            }
            else
            {
                MessageBox.Show("Не вдалося змінити тему інтерфейсу.", "Помилка");
            }
        }

        private void darkThemeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lightToolStripMenuItem.Checked == true)
            {
                this.BackColor = Color.FromArgb(45, 45, 48);
                this.ForeColor = Color.White;

                toolStripTextBox_Filter.BackColor = Color.FromArgb(45, 45, 48);
                toolStripTextBox_Filter.ForeColor = Color.White;

                listView_Service.BackColor = Color.FromArgb(45, 45, 48);
                listView_Service.ForeColor = Color.White;
                listView_Service.Scrollable = true;
                listView_Service.Width += SystemInformation.VerticalScrollBarWidth;
                listView_Service.Width += SystemInformation.HorizontalScrollBarHeight;
                listView_Service.HeaderStyle = ColumnHeaderStyle.None;

                foreach (Control ctrl in this.Controls)
                {
                    ctrl.BackColor = Color.FromArgb(45, 45, 48);
                    ctrl.ForeColor = Color.White;
                    if (ctrl is Button)
                    {
                        Button btn = ctrl as Button;
                        btn.FlatStyle = FlatStyle.Flat;
                        btn.BackColor = Color.FromArgb(28, 28, 28);
                        btn.ForeColor = Color.White;
                    }
                }
                lightToolStripMenuItem.Checked = false;
                darkToolStripMenuItem.Checked = true;
            }
            else
            {
                MessageBox.Show("Не вдалося змінити тему інтерфейсу.", "Помилка");
            }
        }

        private void minFontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChangeFontSize(this.Controls, -1);
        }

        private void standrtFontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetDefaultFontSize(this.Controls, 9);
        }

        private void maxFontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChangeFontSize(this.Controls, 1);
        }

        private void ChangeFontSize(Control.ControlCollection controls, float change)
        {
            foreach (Control ctrl in controls)
            {
                float newSize = ctrl.Font.Size + change;
                if (newSize >= 6)
                {
                    ctrl.Font = new Font(ctrl.Font.FontFamily, newSize, ctrl.Font.Style);
                }

                if (ctrl.Controls.Count > 0)
                {
                    ChangeFontSize(ctrl.Controls, change);
                }
            }
        }

        private void SetDefaultFontSize(Control.ControlCollection controls, float size)
        {
            foreach (Control ctrl in controls)
            {
                ctrl.Font = new Font(ctrl.Font.FontFamily, size, ctrl.Font.Style);

                // Рекурсивно оновити шрифт для вкладених контролів
                if (ctrl.Controls.Count > 0)
                {
                    SetDefaultFontSize(ctrl.Controls, size);
                }
            }
        }

        private void openServicesToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("services.msc");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не вдалося відкрити: " + ex.Message);
            }
        }

        #region 

        [DllImport("gdi32.dll")]
        static extern IntPtr DeleteDC(IntPtr hDc);
        [DllImport("gdi32.dll")]
        static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);
        [DllImport("gdi32.dll")]
        static extern IntPtr CreateCompatibleDC(IntPtr hdc);
        [DllImport("gdi32.dll")]
        static extern IntPtr SelectObject(IntPtr hdc, IntPtr bmp);
        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowDC(IntPtr ptr);
        [DllImport("user32.dll")]
        public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);


        #endregion

    }
}
