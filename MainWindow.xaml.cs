using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Management;
using System.IO;
using System.Media;
using System.Reflection;
using Newtonsoft.Json;
using ModernWPF.Messages;

namespace GameBundler
{
    [Serializable]
    public class CategoryState
    {
        public int CategoryNumber { get; set; }
        public List<string> FilePaths { get; set; } = new List<string>();
        public string Name { get; set; }
    }

    public class Category
    {
        public static int categoryLength = 0;
        public static Category[] arrayOfCategories = new Category[6];

        public int categoryNumber = 0;

        public List<string> FilePaths = new List<string>();

        public List<Process> openedProcesses = new List<Process>();

        public string Name { get; set; }
        public static Grid MainGrid { get; set; }
        public Grid LocalGrid { get; set; } = null;
        public Rectangle Rectangle { get; set; } = null;
        public Label Label { get; set; } = null;

        public Button LaunchBtn { get; set; }
        public Button StopBtn { get; set; }

        public Button AddFilesBtn { get; set; }

        public Image RemoveBtn { get; set; }

        public SoundPlayer Player { get; set; }

        CategoryState state { get; set; } = new CategoryState();

        public Category(Grid mainGrid)
        {
            MainGrid = mainGrid;
            FilePaths = new List<string>();
        }

        public void SavePath()
        {
            if (FilePaths.Count > 0)
            {
                foreach(string path in FilePaths)
                {
                    this.state.FilePaths.Add(path);
                }
            }
        }

        public CategoryState SaveState()
        {

            return new CategoryState
            {
                CategoryNumber = this.categoryNumber,
                FilePaths = this.FilePaths,
                Name = this.Name
            };
        }

        private void SaveCategory(CategoryState categoryState)
        {   
            string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string filePath = System.IO.Path.Combine(docPath + "\\GameBundler", "GameBundler.json");    

            List<CategoryState> categories;

            if (File.Exists(filePath))
            {
                string jsonString = File.ReadAllText(filePath);
                categories = JsonConvert.DeserializeObject<List<CategoryState>>(jsonString);
            }
            else
            {
                categories = new List<CategoryState>();
            }

            int index = categories.FindIndex(c => c.Name == categoryState.Name);

            if (index != -1)
            {
                categories[index] = categoryState;
            }
            else
            {
                categories.Add(categoryState);
            }

            string newJsonString = JsonConvert.SerializeObject(categories);

            File.WriteAllText(filePath, newJsonString);
        }
        public void LoadState(CategoryState state)
        {
            this.categoryNumber = state.CategoryNumber;
            this.FilePaths = state.FilePaths;
            this.Name = state.Name;
        }
        public void RectanglePositionModifier(int catLength)
        {
            Rectangle = new Rectangle();
            Rectangle.Stroke = Brushes.Black;
            Rectangle.Fill = Brushes.Transparent;
            if (catLength == 1) Rectangle.Margin = new Thickness(30, 70, 690, 310);
            else if (catLength == 2) Rectangle.Margin = new Thickness(360, 70, 360, 310);
            else if (catLength == 3) Rectangle.Margin = new Thickness(690, 65, 30, 310);
            else if (catLength == 4) Rectangle.Margin = new Thickness(30, 281, 690, 100);
            else if (catLength == 5) Rectangle.Margin = new Thickness(360, 281, 360, 100);
            else if (catLength == 6) Rectangle.Margin = new Thickness(690, 281, 30, 100);
            else Rectangle = null;
        }


        public void RemoveBtnPositionModifier(int catLength)
        {
            RemoveBtn = new Image();
            RemoveBtn.Source = new BitmapImage(new Uri("pack://application:,,,/GameBundler;component/Resources/RemoveIcon.png"));
            RemoveBtn.MouseLeftButtonDown += (sender, e) => RemoveBtnAction(sender, e, Name);
            if (catLength == 1) RemoveBtn.Margin = new Thickness(274, 76, 696, 470);
            else if (catLength == 2) RemoveBtn.Margin = new Thickness(604, 76, 366, 470);
            else if (catLength == 3) RemoveBtn.Margin = new Thickness(934, 76, 36, 470);
            else if (catLength == 4) RemoveBtn.Margin = new Thickness(274, 290, 696, 256);
            else if (catLength == 5) RemoveBtn.Margin = new Thickness(604, 290, 366, 256);
            else if (catLength == 6) RemoveBtn.Margin = new Thickness(934, 290, 36, 256); 


            else RemoveBtn = null;
        }

        public void OtherBtnsModifier(int catLength)
        {
            AddFilesBtn = new Button(); AddFilesBtn.Visibility = Visibility.Collapsed; AddFilesBtn.Content = "Add Files"; AddFilesBtn.Foreground = Brushes.White;
            LaunchBtn = new Button(); LaunchBtn.Visibility = Visibility.Collapsed; LaunchBtn.Content = "Launch Category"; LaunchBtn.Foreground = Brushes.White;
            StopBtn = new Button(); StopBtn.Visibility = Visibility.Collapsed; StopBtn.Content = "Stop Category"; StopBtn.Foreground = Brushes.White;

            Style style = Application.Current.TryFindResource("RoundedButton") as Style;
            Brush hotTrackBrush = Application.Current.FindResource(SystemColors.HotTrackBrushKey) as Brush;
            Brush WindowFrameBrushKey = Application.Current.TryFindResource("WindowFrameBrushKey") as Brush;

            AddFilesBtn.Style = style; LaunchBtn.Style = style; StopBtn.Style = style;

            AddFilesBtn.Background = WindowFrameBrushKey;
            LaunchBtn.Background = hotTrackBrush; StopBtn.Background = Brushes.Red;


            AddFilesBtn.Click += (s, e) =>
            {
                AddFileFunc(); 
            };

            LaunchBtn.Click += (s, e) =>
            {
                LaunchCategoryFunc();
            };

            StopBtn.Click += (s, e) =>
            {
                StopCategoryFunc();
            };

            if (catLength == 1)
            {
                AddFilesBtn.Margin = new Thickness(95, 127, 750, 396);
                LaunchBtn.Margin = new Thickness(95, 183, 750, 340);
                StopBtn.Margin = new Thickness(95, 183, 750, 340);
            }
            else if (catLength == 2)
            {
                AddFilesBtn.Margin = new Thickness(418, 127, 427, 396);
                LaunchBtn.Margin = new Thickness(418, 183, 427, 340);
                StopBtn.Margin = new Thickness(418, 183, 427, 340);
            }
            else if (catLength == 3)
            {
                AddFilesBtn.Margin = new Thickness(756, 128, 88, 396);
                LaunchBtn.Margin = new Thickness(757, 183, 88, 340);
                StopBtn.Margin = new Thickness(757, 183, 88, 340);
            }
            else if (catLength == 4)
            {
                AddFilesBtn.Margin = new Thickness(95, 338, 750, 184);
                LaunchBtn.Margin = new Thickness(95, 393, 750, 130);
                StopBtn.Margin = new Thickness(95, 393, 750, 130);
            }
            else if (catLength == 5)
            {
                AddFilesBtn.Margin = new Thickness(418, 339, 427, 184);
                LaunchBtn.Margin = new Thickness(418, 393, 427, 130);
                StopBtn.Margin = new Thickness(418, 393, 427, 130);
            }
            else if (catLength == 6)
            {
                AddFilesBtn.Margin = new Thickness(757, 340, 87, 184);
                LaunchBtn.Margin = new Thickness(757, 340, 87, 184);
                StopBtn.Margin = new Thickness(757, 340, 87, 184);
            }

        }

        public void ShowOtherBtns()
        {
            AddFilesBtn.Visibility = Visibility.Visible;
            LaunchBtn.Visibility = Visibility.Visible;
        }
        public void HideOtherBtns()
        {
            AddFilesBtn.Visibility = Visibility.Collapsed;
            LaunchBtn.Visibility = Visibility.Collapsed;
        }

        private void RemoveBtnAction(object sender, MouseButtonEventArgs e, string catName)
        {
            int catNumber = Array.FindIndex(arrayOfCategories, cat => cat != null && cat.Name == catName);
            if (catNumber < 0)
            {
                return;
            }

            arrayOfCategories[catNumber] = null;
            categoryLength--;
            UpdateCategory();

            string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string filePath = System.IO.Path.Combine(docPath + "\\GameBundler", "GameBundler.json");

            List<CategoryState> categories;

            if (File.Exists(filePath))
            {
                string jsonString = File.ReadAllText(filePath);
                categories = JsonConvert.DeserializeObject<List<CategoryState>>(jsonString);

                int index = categories.FindIndex(cat => cat.Name == catName);
                if (index >= 0)
                {
                    categories.RemoveAt(index);
                }

                string newJsonString = JsonConvert.SerializeObject(categories);

                File.WriteAllText(filePath, newJsonString);
            }
        }


        public void AddBlurOnHovering(Rectangle rectangle, Label label, Grid localGrid)
        {
            BlurEffect blur = new BlurEffect();
            blur.Radius = 0;
            rectangle.Effect = blur;
            label.Effect = blur;
            localGrid.MouseEnter += (s, e) =>
            {
                blur.Radius = 5;
                ShowOtherBtns();
            };
            localGrid.MouseLeave += (s, e) =>
            {
                blur.Radius = 0;
                HideOtherBtns();
            };
        }

        public void AddFileFunc()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                FilePaths.Add(openFileDialog.FileName);
                state.FilePaths.Add(openFileDialog.FileName);
                Player = new SoundPlayer();
                var assembly = Assembly.GetExecutingAssembly();
                var stream = assembly.GetManifestResourceStream("GameBundler.success.wav");
                Player.Stream = stream;
                Player.Play();
            }
            SaveCategory(SaveState());
        }

        public void LaunchCategoryFunc()
        {
            if (FilePaths.Count <= 0)
            {
                MessageBox.Show("Please Add Files First!");
                return;
            }
            foreach (string appPath in FilePaths)
            {
                Process process = Process.Start(appPath);
                openedProcesses.Add(process);
            }
            LaunchBtn.Visibility = Visibility.Collapsed;
            StopBtn.Visibility = Visibility.Visible;
        }

        public void StopCategoryFunc()
        {
            foreach (Process process in openedProcesses)
            {
                if (process != null)
                {
                    if(process.ProcessName.Length <= 3)
                    {
                        var proc = Process.GetProcesses().Where(p => p.ProcessName.Contains(process.ProcessName.Substring(0, 3)));
                        foreach (Process p in proc) p.Kill();
                    }
                    else
                    {
                        var proc = Process.GetProcesses().Where(p => p.ProcessName.Contains(process.ProcessName.Substring(0, 5)));
                        foreach (Process p in proc) p.Kill();
                    }

                }
            }
            foreach (Process process in openedProcesses)
            {
                try
                {
                    if (process != null )
                    {
                        List<int> childProcessIds = GetChildProcesses(process.Id);

                        foreach (int childPid in childProcessIds)
                        {
                            KillProcessAndChildren(childPid);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to kill process: {ex.Message}");
                }
            }
            openedProcesses.Clear();
            LaunchBtn.Visibility = Visibility.Visible;
            StopBtn.Visibility = Visibility.Collapsed;
        }

        public void KillProcessAndChildren(int pid)
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher
                ("Select * From Win32_Process Where ParentProcessID=" + pid);
            ManagementObjectCollection moc = searcher.Get();

            foreach (ManagementObject mo in moc)
            {
                KillProcessAndChildren(Convert.ToInt32(mo["ProcessID"]));
            }

            try
            {
                Process proc = Process.GetProcessById(pid);
                proc.Kill();
            }
            catch (ArgumentException)
            {
            }
        }

        public List<int> GetChildProcesses(int parentProcessId)
        {
            List<int> childProcessIds = new List<int>();
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(
                "SELECT * " +
                "FROM Win32_Process " +
                "WHERE ParentProcessId=" + parentProcessId);
            ManagementObjectCollection collection = searcher.Get();
            foreach (var item in collection)
            {
                childProcessIds.Add(Convert.ToInt32(item["ProcessId"]));
            }
            return childProcessIds;
        }


        public int CategoryNumberFinder(Category[] array)
        {
            for (int i = 0; i < array.Length; i++)
            {
                if (array[i] == null) return i + 1;
            }
            return -1;
        }

        public void AddNewCategory(string catName)
        {
            Name = catName;
            categoryLength++;
            categoryNumber = CategoryNumberFinder(arrayOfCategories);
            LocalGrid = new Grid();
      
            RectanglePositionModifier(this.categoryNumber);
            RemoveBtnPositionModifier(this.categoryNumber);
            OtherBtnsModifier(this.categoryNumber);

            Label = new Label();
            Label.Foreground = Brushes.White;
            Label.Height = 31;
            Label.Width = Rectangle.Width;
            Label.FontSize = 16;
            Label.FontWeight = FontWeights.Bold;
            Label.Margin = Rectangle.Margin;
            Label.VerticalContentAlignment = VerticalAlignment.Center;
            Label.HorizontalContentAlignment = HorizontalAlignment.Center;
            Label.VerticalAlignment = VerticalAlignment.Top;
            Label.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#63216B7F"));
            Label.Content = catName;

            AddBlurOnHovering(Rectangle, Label, LocalGrid);

            LocalGrid.Children.Add(Label);
            LocalGrid.Children.Add(Rectangle);
            LocalGrid.Children.Add(RemoveBtn);
            LocalGrid.Children.Add(AddFilesBtn);
            LocalGrid.Children.Add(LaunchBtn);
            LocalGrid.Children.Add(StopBtn);
            arrayOfCategories[categoryNumber - 1] = this;
        }
        public static void UpdateCategory()
        {
            List<Grid> gridsToRemove = new List<Grid>();
            foreach (UIElement child in MainGrid.Children)
            {
                if (child is Grid)
                {
                    gridsToRemove.Add((Grid)child);
                }
            }
            foreach (Grid grid in gridsToRemove)
            {
                MainGrid.Children.Remove(grid);
            }

            foreach (Category cat in arrayOfCategories)
            {
                if (cat != null && !MainGrid.Children.Contains(cat.LocalGrid))
                {

                    MainGrid.Children.Add(cat.LocalGrid);
                }
            }
        }
    }
    public partial class MainWindow : Window
    {
        private System.Windows.Forms.NotifyIcon notifyIcon = null;
        public MainWindow()
        {
            InitializeComponent();
            LoadCategories();
            this.MouseLeftButtonDown += MainWindow_MouseLeftButtonDown;
            this.notifyIcon = new System.Windows.Forms.NotifyIcon();
            this.notifyIcon.Icon = new System.Drawing.Icon(Properties.Resources.MainIcon, new System.Drawing.Size(16,16));
            this.notifyIcon.Visible = false;

            System.Windows.Forms.ContextMenuStrip contextMenu = new System.Windows.Forms.ContextMenuStrip();
            System.Windows.Forms.ToolStripMenuItem closeMenuItem = new System.Windows.Forms.ToolStripMenuItem("Close");
            closeMenuItem.Click += CloseMenuItem_Click;
            contextMenu.Items.Add(closeMenuItem);
            this.notifyIcon.ContextMenuStrip = contextMenu;

            this.notifyIcon.DoubleClick += (s, args) =>
            {
                this.Show();
                WindowState = WindowState.Normal;
                this.notifyIcon.Visible = false;
            };

        }



        void MainWindow_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is Image)
            {
                return;
            }
            DragMove();
        }

        private void CloseMenuItem_Click(object sender, EventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void Ellipse_MouseEnter(object sender, MouseEventArgs e)
        {
            Ellipse ellipse = sender as Ellipse;
            if (ellipse != null)
            {
                ellipse.Fill = Brushes.Green; 
            }
        }

        private void Ellipse_MouseLeave(object sender, MouseEventArgs e)
        {
            Ellipse ellipse = sender as Ellipse;
            if (ellipse != null)
            {
                if (ellipse.Name == "CloseEllipse")
                {
                    ellipse.Fill = Brushes.Red;
                }
                else if (ellipse.Name == "MinimizeEllipse")
                {
                    ellipse.Fill = Brushes.Yellow;
                }
            }
        }

        private void MinimizeToTrayButton_Click(object sender, RoutedEventArgs e)
        {
            // Minimize window and show tray icon
            this.WindowState = WindowState.Minimized;
            this.Hide();
            this.notifyIcon.Visible = true;
        }

        private void Minimize(object sender, RoutedEventArgs e)
        {
            // Minimize window and show tray icon
            this.WindowState = WindowState.Minimized;
        }


        private void BlurWindow()
        {
            var blur = new System.Windows.Media.Effects.BlurEffect();
            blur.Radius = 5;
            this.Effect = blur;
        }

        private void UnblurWindow()
        {
            this.Effect = null;
        }

        private void AddCategory_Click(object sender, RoutedEventArgs e)
        {
            if (Category.categoryLength >= 6)
            {
                MessageBox.Show("Categories are FULL! Please remove some categories.");
                return;
            }
            CustomDialog dialog = new CustomDialog();
            dialog.Owner = this;
            dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            BlurWindow();
            dialog.ShowDialog();
            UnblurWindow();
            if (dialog.Confirmed)
            {
                AddNewCategory(dialog.CategoryName, new List<string>(), true);
            }
        }

        private void AddNewCategory(string catName,List<string> filePath, bool saveToFile = true)
        {
            Category newCategory = new Category(MainGrid);
            newCategory.FilePaths = filePath;
            newCategory.AddNewCategory(catName);
            Category.UpdateCategory();
            if (saveToFile)
            {
                SaveCategory(newCategory.SaveState());
            }


        }
        private void SaveCategory(CategoryState categoryState)
        {
            string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string filePath = System.IO.Path.Combine(docPath + "\\GameBundler", "GameBundler.json");

            List<CategoryState> categories;

            if (File.Exists(filePath))
            {
                string jsonString = File.ReadAllText(filePath);
                categories = JsonConvert.DeserializeObject<List<CategoryState>>(jsonString);
            }
            else
            {
                categories = new List<CategoryState>();
            }

            categories.Add(categoryState);

            string newJsonString = JsonConvert.SerializeObject(categories);

            File.WriteAllText(filePath, newJsonString);
        }

        public void LoadCategories()
        {
            
            string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string dirPath = System.IO.Path.Combine(docPath, "GameBundler");
            string filePath = System.IO.Path.Combine(dirPath, "GameBundler.json");
            string jsonString;

            if (!Directory.Exists(dirPath))
            {
                Directory.CreateDirectory(dirPath);
            }

            if (!File.Exists(filePath))
            {
                List<CategoryState> emptyList = new List<CategoryState>();
                jsonString = JsonConvert.SerializeObject(emptyList);
                File.WriteAllText(filePath, jsonString);
            }

            jsonString = File.ReadAllText(filePath);
            List<CategoryState> categoryStates = JsonConvert.DeserializeObject<List<CategoryState>>(jsonString);

            foreach (CategoryState categoryState in categoryStates)
            {
                Category newCategory = new Category(MainGrid);
                newCategory.LoadState(categoryState);
                AddNewCategory(newCategory.Name,newCategory.FilePaths, false);
            }
        }

    }

}
