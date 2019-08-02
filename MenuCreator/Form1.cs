using IniParser;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Renci.SshNet;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using WMPLib;
using AppForm = System.Windows.Forms.Application;
using Excel = Microsoft.Office.Interop.Excel;

namespace MenuCreator
{
    public partial class Form1 : Form
    {
        public Form1(string[] Args)
        {
            InitializeComponent();
            if (Args.Length > 0) {
                Console.WriteLine("Arguments found, dealing with the arguie");
                Console.WriteLine(Args);
                shouldChange = true;
                this.Args = Args;
            }
        }

        public bool shouldChange = false;
        public string[] Args = { };

        //Open Excel File
        private void button1_Click(object sender, EventArgs e)
        {
            DisposePictureBox();
            if (radioButton1.Checked)
            {
                Process.Start(AppForm.StartupPath + @"\StrainMenuCreator.exe");
                this.Close();
            }
            else
            {
                try {
                    string Menu = GetMenuString() + ".xlsx";
                    Process.Start(Menu);
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            }
        }

        //Create Image
        public string FileLocation;

        public string XLSX;
        public string FileName;
        public string Output;

        private string GetMenuString()
        {
            string Menu = "";
            if (jointRadio.Checked)
                Menu = "Joint_Menu";
            else if (edibleRadio.Checked)
                Menu = "Edible_Menu";
            else if (cartRadio.Checked)
                Menu = "Cart_Menu";
            else if (dabRadio.Checked)
                Menu = "Dab_Menu";
            else if (dailyRadio.Checked)
                Menu = "Daily_Special";
            else
                Menu = "Menu";
            return Menu;
        }

        public bool Export_1080p = false;
        public string webURL = "http://192.168.1.210/manage/shares/Server/Menus/MenuCreator/Uploads/";

        private void button2_Click(object sender, EventArgs e)
        {
            Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\MenuImages\\");
            try { CreateMenu(); }
            catch
            {
                Console.WriteLine("Menu creation failed, trying again....");
                DisposePictureBox();
                File.Delete("Menu_Small.png");
                FileName = GetMenuString();
                Output = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\MenuImages\\" + FileName + "_" + DateTime.Today.ToString("MM-dd-yyyy") + "_.png";
                File.Delete(Output);
                CreateMenu();
            }

            if (dailyRadio.Checked)
            {
                DisposePictureBox();
                CreateDailyMenu();
            }

            if (autoUpload.Checked)
            {
                webURL = @"/home/pi/screenly_assets/AUTOMATED_" + DateTime.Now.ToString("MM-dd-yyyy_hhmm");
                string Output =
                    Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) +
                    "\\MenuImages\\" + GetMenuString() + "_" + DateTime.Today.ToString("MM-dd-yyyy") + "_.png";

                Console.WriteLine(Output);
                Console.WriteLine(webURL);
                Console.WriteLine("Deleting all old assets");
                DeleteOldAssetsAsync(GetIP());
                Console.WriteLine("Uploading to Screenly Menus");
                DisposePictureBox();

                SFTPUpload(Output, webURL, GetIP());
                Upload(webURL, GetIP());
            }
        }

        public async void DeleteOldAssetsAsync(string IP)
        {
            Device newDevice = new Device
            {
                Name = "Specials",
                Location = "Floor",
                IpAddress = IP,
                Port = "80",
                ApiVersion = "v1.1/"
            };

            await newDevice.GetAssetsAsync();
            foreach (Asset asset in newDevice.ActiveAssets)
            {
                Console.WriteLine(asset.AssetId);
                await newDevice.RemoveAssetAsync(asset.AssetId);
            }
        }

        public void CreateDailyMenu()
        {
            Cursor.Current = Cursors.WaitCursor;
            cmdOutput = new StringBuilder("");
            string Desktop = Environment.GetFolderPath
                (Environment.SpecialFolder.DesktopDirectory);
            string ImageMagick = @"C:\Program Files\ImageMagick";

            int X = int.Parse(xPos.Text);
            int Y = int.Parse(yPos.Text);

            string st = AppForm.StartupPath + "\\";

            Console.WriteLine("Making Daily Menu All Pretty");
            FileName = GetMenuString();
            Output = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\MenuImages\\" + FileName + "_" + DateTime.Today.ToString("MM-dd-yyyy") + "_.png";
            string img = "Daily_Image.png";
            File.Delete(img);
            File.Delete("cut_" + img);
            File.Move(Output, img);
            string Command2 = String.Format(@"cd {1} & magick convert {0} +repage -shave 1x1 {0}", st + img, ImageMagick);
            cmd(Command2, true, true, true);

            File.Delete(Output);
            File.Copy("Daily_Template.png", Output);

            string e = (Y < 0) ? "" : "+";
            string f = (X < 0) ? "" : "+";

            Console.WriteLine("Editing: Daily_Image.png");
            string Command = String.Format(@"cd {4} & magick convert {1} {0} -gravity center -geometry {6}{2}{5}{3} -composite {1}", st + img, Output, X, Y, ImageMagick, e, f);
            Console.WriteLine(Command);
            cmd(Command, true, true, true);
            Cursor.Current = Cursors.Default;

            PreviewImage();
            MessageBox.Show("Done, image should be located on your desktop!");
        }

        public void CreateMenu()
        {
            Cursor.Current = Cursors.WaitCursor;
            Clipboard.Clear();
            foreach (var process in Process.GetProcessesByName("EXCEL"))
            {
                process.Kill();
                process.WaitForExit();
            }

            DisposePictureBox();
            FileName = GetMenuString();
            XLSX = FileName + ".xlsx";
            Output = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\MenuImages\\" + FileName + "_" + DateTime.Today.ToString("MM-dd-yyyy") + "_.png";
            FileLocation = AppForm.StartupPath + "\\" + XLSX;
            File.Delete(Output);
            if (File.Exists(XLSX)) {
                ConvertFile(Output);
                PreviewImage();
            }
            else
                if (!dailyRadio.Checked)
                MessageBox.Show("Excel File is missing!");
            Cursor.Current = Cursors.Default;
        }

        public void ConvertFile(string output)
        {
            Clipboard.Clear();
            CreateImage_Alt(excelRange1.Text, excelRange2.Text);
            if (!dailyRadio.Checked)
                ResizeImage("Menu_Small.png");
            else
                File.Move("Menu_Small.png", output);

            if (!autoUpload.Checked)
            {
                if (File.Exists(output) && !dailyRadio.Checked)
                    MessageBox.Show("Image Created! It should be located on your desktop.");
                else
                    Console.WriteLine("Image creation failed....");
            }
        }

        public string ReadINI(string Key = "Settings", string Object = "")
        {
            var parser = new FileIniDataParser();
            var data = parser.ReadFile("Settings.ini");
            return data[Key][Object];
        }

        public void CreateImage_Alt(string Range1, string Range2)
        {
            Console.WriteLine("Creating initial image...");
            Excel.Application excel = new Excel.Application();
            Workbook w = excel.Workbooks.Open(FileLocation);
            Worksheet ws = w.Sheets[1];
            ws.Protect(Contents: false);
            Console.WriteLine("Range is " + Range1 + ":" + Range2);
            string ImageRange = Range1 + ":" + Range2;
            Range r = ws.Range[ImageRange];
            r.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);
            try
            {
                Bitmap image = new Bitmap(Clipboard.GetImage());
                image.Save("Menu_Small.png");
                Console.WriteLine("Small image saved...");
                image.Dispose();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error!");
                Console.WriteLine(e.ToString());
                Console.ReadLine();
            }
            w.Close(false, Type.Missing, Type.Missing);
            excel.Quit();
        }

        public void DisposePictureBox()
        {
            if (pictureBox1.Image != null)
                try
                {
                    Console.WriteLine("Disposing picture box image...");
                    var image = pictureBox1.Image;
                    pictureBox1.Image = null;
                    image.Dispose();
                }
                catch { Console.WriteLine("Failed to dispose of picture box image, probably due to picture box not having an image."); }
        }

        public void ResizeImage(string fileName)
        {
            Console.WriteLine("Resizing image...");
            FileInfo info = new FileInfo(fileName);
            using (Image image = Image.FromFile(fileName))
            {
                int Width = 1;
                int Height = 1;

                if (radio_4k.Checked)
                {
                    Width = 3840;
                    Height = 2160;
                }
                else if (radio_1080.Checked)
                {
                    Width = 1920;
                    Height = 1080;
                }
                else
                {
                    Width = Int32.Parse(cSizeW.Text);
                    Height = Int32.Parse(cSizeH.Text);
                }

                using (Bitmap resizedImage = ResizeImage(image, Width, Height))
                    resizedImage.Save(Output);
            }
            Console.WriteLine("Done");
            File.Delete("Menu_Small.png");
        }

        private void Extract(string nameSpace, string outDirectory, string internalFilePath, string resourceName)
        {
            var assembly = Assembly.GetCallingAssembly();
            using (var s = assembly.GetManifestResourceStream(nameSpace + "." +
                                                   (internalFilePath == "" ? "" : internalFilePath + ".") +
                                                   resourceName))
            using (var r = new BinaryReader(s))
            using (var fs = new FileStream(outDirectory + "\\" + resourceName, FileMode.OpenOrCreate))
            using (var w = new BinaryWriter(fs))
                w.Write(r.ReadBytes((int)s.Length));
        }

        public Bitmap ResizeImage(Image image, int width, int height)
        {
            Console.WriteLine("Width: " + width);
            Console.WriteLine("Height: " + height);
            var destRect = new System.Drawing.Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }

        public Bitmap ResizeImage(Image image, Decimal Width, Decimal Height)
        {
            return ResizeImage(image, Width, Height);
        }

        //Preview Button
        private void button3_Click(object sender, EventArgs e)
        {
            PreviewImage();
        }

        private void PreviewImage()
        {
            DisposePictureBox();
            try
            {
                string ImageLocation = Environment.GetFolderPath
                                           (Environment.SpecialFolder.DesktopDirectory) + "\\MenuImages\\" + GetMenuString() + "_" +
                                       DateTime.Today.ToString("MM-dd-yyyy") + "_.png";
                pictureBox1.Image = Image.FromFile(ImageLocation);
            }
            catch { }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Process.Start(@"Z:\Menus - OLD\Menu Manager\index.html");
        }

        private string GetIP()
        {
            if (jointRadio.Checked)
                return "192.168.1.111";
            else if (edibleRadio.Checked)
                return "192.168.1.114";
            else if (cartRadio.Checked)
                return "192.168.1.115";
            else if (dabRadio.Checked)
                return "192.168.1.116";
            else if (radioButton1.Checked)
                return "192.168.1.112";
            else if (dailyRadio.Checked)
                return "192.168.1.67";
            else
                return "";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
#if !DEBUG
            PingAll();
            StartTimer();
#endif
            if (Process.GetProcessesByName("MenuCreator").Count() > 1)
                this.Close();
            if (!File.Exists("Settings.ini"))
                Extract("MenuCreator", System.Windows.Forms.Application.StartupPath, "Files", "Settings.ini");

            File.Delete("Menu_Small.png");

            _writer = new TextBoxStreamWriter(txtConsole);
            Console.SetOut(_writer);

            xPos.Text = ReadINI("Settings", "xPos");
            yPos.Text = ReadINI("Settings", "yPos");
            fDelay.Text = ReadINI("Settings", "fDelay");

            Process p = new Process();
            p.StartInfo.FileName = "explorer.exe";
            p.StartInfo.Arguments = @"\\192.168.1.210\Server\Menus\";
            p.Start();
            p.Kill();
            p.Dispose();

            Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\MenuImages\\");
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            button2.Enabled = !radioButton1.Checked;
            button5.Enabled = radioButton1.Checked;
            button7.Enabled = radioButton1.Checked;
            button8.Enabled = radioButton1.Checked;
            groupBox3.Enabled = !radioButton1.Checked;

            if (radioButton1.Checked)
            {
                groupBox2.Enabled = true;
                fDelay.Enabled = true;
                noExtract.Enabled = true;
            }
            else if (dailyRadio.Checked)
            {
                groupBox2.Enabled = true;
                fDelay.Enabled = false;
                noExtract.Enabled = false;
            }
            else
                groupBox2.Enabled = false;

        }

        /*
         * Uploads 'movie.mp4' to the Flower Menu Raspberry Pi and plays it.
         * Done via ssh.
         * Typically uploads a 5MB video in a second or two, but results may differ on your internet
         */

        private void button5_Click(object sender, EventArgs e)
        {
            if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\MenuImages\\movie.mp4"))
            {
                Cursor.Current = Cursors.WaitCursor;
                try
                {
                    Console.WriteLine("Uploading movie.mp4");
                    SFTPUpload(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\MenuImages\\movie.mp4", @"/home/pi/screenly_assets/movie.mp4");
                    Console.WriteLine("Done");
                }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                Cursor.Current = Cursors.Default;
            }
            else
                MessageBox.Show("There's no movie named 'movie.mp4' on your desktop. If you don't know what you're doing you shouldn't be running this command.");
        }

        public void SFTPUpload(string fileToUpload, string fileLocation, string host = "192.168.1.112", string user = "pi", string password = "raspberry", int Port = 22)
        {
            try
            {
                var client = new SftpClient(host, Port, user, password);
                client.Connect();
                if (client.IsConnected)
                    using (var fileStream = new FileStream(fileToUpload, FileMode.Open))
                    {
                        client.UploadFile(fileStream, fileLocation);
                        client.Disconnect();
                        client.Dispose();
                    }
                else
                    Console.WriteLine("Couldn't connect to host");
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
        }

        public void SSH(string Command, string IP = "192.168.1.112")
        {
            using (var client = new SshClient(IP, "pi", "raspberry"))
            {
                client.Connect();
                client.RunCommand(Command);
                client.Disconnect();
            }
        }

        private void ScpUpload(string filePath, string destinationFilePath, string host = "192.168.1.112", int port = 22, string username = "pi", string password = "raspberry")
        {
            ConnectionInfo connInfo = new ConnectionInfo(host, username, new PasswordAuthenticationMethod(username, password));
            using (var scp = new ScpClient(connInfo))
            {
                scp.Connect();
                scp.Upload(new FileInfo(filePath), destinationFilePath);
                scp.Disconnect();
            }
            Console.WriteLine("Upload function done");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SSH("sudo reboot", GetIP());
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var task = Task.Run(() => {
                SSH(@"sudo kill -9 `pgrep omxplayer` & echo done", "192.168.1.112"); });
            bool isCompletedSuccessfully = task.Wait(TimeSpan.FromMilliseconds(10000));
        }

        public void cmd(string Arguments, bool isHidden = false, bool waitForexit = true, bool redirect = true)
        {
            ProcessStartInfo ProcessInfo;
            Process Process;
            try
            {
                ProcessInfo = new ProcessStartInfo(@"C:\Windows\system32\cmd.exe", "/C " + Arguments);
                ProcessInfo.UseShellExecute = false;
                ProcessInfo.CreateNoWindow = isHidden;
                ProcessInfo.RedirectStandardOutput = redirect;
                Process = Process.Start(ProcessInfo);
                Process.BeginOutputReadLine();
                Process.OutputDataReceived += new DataReceivedEventHandler(SortOutputHandler);
                Process.ErrorDataReceived += new DataReceivedEventHandler(SortOutputHandler);

                if (waitForexit == true)
                    Process.WaitForExit();

                Process.Close();
                Process.Dispose();
                Process = null;
            }
            catch (Exception e)
            {
                Console.WriteLine("Error Processing ExecuteCommand : " + e.Message);
            }
        }

        private static void Process_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            Console.WriteLine(e.Data);
        }

        private static void SortOutputHandler(object sendingProcess, DataReceivedEventArgs outLine)
        {
            if (!String.IsNullOrEmpty(outLine.Data))
                cmdOutput.Append(Environment.NewLine + outLine.Data);
        }

        private bool BootTooBig()
        {
            string fileLocation = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\MenuImages\\Menu_" + DateTime.Today.ToString("MM-dd-yyyy") + "_.png";
            using (Image img = Image.FromFile(fileLocation))
            {
                if (img.Width > 1920 || img.Height > 1080)
                    return true;
                else
                    return false;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            Console.WriteLine("Starting Video creation process. This can take upwards of 10 minutes depending on your computer, GIF size/length and image size.");
            if (!BootTooBig())
                VideoCreation();
            else
            {
                DialogResult dialogResult = MessageBox.Show("Your image is larger than 1080p, do you want to continue making this into a video? Doing so can make the process take considerablly longer or even cause it to fail if your computer cant handle larger resolutions.", "Export options", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                    VideoCreation();
                else
                    Console.WriteLine("Video Creation cancelled");
            }
            Cursor.Current = Cursors.Default;
        }

        private void VideoCreation()
        {
            Cursor.Current = Cursors.WaitCursor;
            cmdOutput = new StringBuilder("");
            string Desktop = Environment.GetFolderPath
                (Environment.SpecialFolder.DesktopDirectory);
            string ImageMagick = @"C:\Program Files\ImageMagick";
            Output = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\MenuImages\\Menu_" + DateTime.Today.ToString("MM-dd-yyyy") + "_.png";
            if (File.Exists(Output) && File.Exists(Path.Combine(Desktop, "\\MenuImages\\Insert.gif")))
            {
                String Command = "";
                if (!noExtract.Checked)
                {
                    try
                    {
                        Directory.Delete(Path.Combine(Desktop, "Gif"), true);
                        Console.WriteLine("Deleted old GIF");
                    }
                    catch { }
                    Console.WriteLine("Both the GIF and menu are located extracting gif");
                    Directory.CreateDirectory(Path.Combine(Desktop, "Gif"));
                    Command = @"cd " + ImageMagick + " & magick convert -coalesce " + Path.Combine(Desktop, "\\MenuImages\\Insert.gif") + " " + Path.Combine(Desktop, "Gif\\Target.png");
                    cmd(Command, true, true, true);
                    Console.WriteLine("GIF Extracted");
                    int i = 0;
                    string Target = Path.Combine(Desktop, "Gif");
                    while (i < 10)
                    {
                        Target = Path.Combine(Desktop, "Gif");
                        string img = "Target-" + i + ".png";
                        if (File.Exists(Path.Combine(Target, img)))
                            File.Move(Path.Combine(Target, img), Path.Combine(Target, "Target-0" + i.ToString() + ".png"));
                        i++;
                    }

                    foreach (string img in Directory.GetFiles(Path.Combine(Desktop, "Gif"), "*.png"))
                    {
                        string X = xPos.Text;
                        string Y = yPos.Text;
                        Console.WriteLine("Editing: " + img);
                        Command = @"cd " + ImageMagick + " & magick convert " + Output + " " + img +
                                  " -gravity southeast -geometry +" + X + "+" + Y + " -composite " + img;
                        cmd(Command, true, true, true);
                    }
                }

                Console.WriteLine("Recreating GIF");
                Command = @"cd " + ImageMagick + " & magick convert -delay " + fDelay.Text + " -loop 0 " + Path.Combine(Desktop, "Gif\\Target-*.png") + " C:\\Users\\Joe\\Desktop\\MenuImages\\Menu_GIF.gif";
                cmd(Command, true, true, true);

                Console.WriteLine(Path.Combine(Desktop, "\\MenuImages\\Output_Gif_" + DateTime.Today.ToString("MM-dd-yyyy")) + ".gif");
                Console.WriteLine("Creating mp4 file out of GIF");
                Command = AppForm.StartupPath + "\\HandBrakeCLI.exe -Z \"Very Fast 1080p30\" -i C:\\Users\\Joe\\Desktop\\MenuImages\\Menu_GIF.gif -o " + Path.Combine(Desktop, "MenuImages\\movie.mp4") + " & echo done";
                cmd(Command, true, true, true);

                try {
                    if (Duration(Path.Combine(Desktop, "\\MenuImages\\movie.mp4")) < 2)
                    {
                        Console.WriteLine("Video is 2 seconds or under extending...");
                        CombineVideos();
                        Console.WriteLine("Done");
                    }
                }
                catch (Exception ex) {
                    Console.WriteLine("Error on combining videos!");
                    Console.WriteLine(ex.ToString());
                }

                try { Directory.Delete(Path.Combine(Desktop, "Gif"), true); }
                catch { }
            }
            else
                MessageBox.Show("The menu needs to be on your desktop exactly as it was when it was created, and your gif needs to be on your desktop named 'insert.gif'.");

            Cursor.Current = Cursors.Default;
        }

        public void CombineVideos()
        {
            string Desktop = Environment.GetFolderPath
                (Environment.SpecialFolder.DesktopDirectory);
            File.Delete(Path.Combine(AppForm.StartupPath, "movie_combined.mp4"));
            File.Move(Path.Combine(Desktop, "\\MenuImages\\movie.mp4"), AppForm.StartupPath + "\\movie.mp4");

            string Commands = "-safe 0 -f concat -i list.txt -c copy movie_combined.mp4";
            cmdOutput = new StringBuilder("");
            Process p = new Process();
            p.StartInfo.FileName = AppForm.StartupPath + "\\ffmpeg.exe";
            p.StartInfo.Arguments = Commands;
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.CreateNoWindow = false;
            p.StartInfo.RedirectStandardOutput = false;
            p.Start();
            p.WaitForExit();
            p.Dispose();
            p = null;

            File.Delete(Path.Combine(Desktop, "\\MenuImages\\movie.mp4"));
            File.Move(Path.Combine(AppForm.StartupPath, "movie_combined.mp4"), Path.Combine(Desktop, "\\MenuImages\\movie.mp4"));
        }

        public Double Duration(String file)
        {
            WindowsMediaPlayer wmp = new WindowsMediaPlayer();
            IWMPMedia mediainfo = wmp.newMedia(file);
            Console.WriteLine("Duration: " + mediainfo.duration);
            return mediainfo.duration;
        }

        private static StringBuilder cmdOutput = null;
        private TextWriter _writer = null;

        public class TextBoxStreamWriter : TextWriter
        {
            private System.Windows.Forms.TextBox _output = null;

            public TextBoxStreamWriter(System.Windows.Forms.TextBox output)
            {
                _output = output;
            }

            public override void Write(char value)
            {
                try
                {
                    base.Write(value);
                    _output.AppendText(value.ToString()); // When character data is written, append it to the text box.
                }
                catch { }
            }

            public override Encoding Encoding
            {
                get { return System.Text.Encoding.UTF8; }
            }
        }

        private void xPos_TextChanged(object sender, EventArgs e)
        {
            if (!dailyRadio.Checked)
                Write("xPos", xPos.Text);
            else
            {
                var parser = new FileIniDataParser();
                var data = parser.ReadFile("Settings.ini");

                data["Daily_Special"]["xPos"] = xPos.Text ?? "NA";
                parser.WriteFile("Settings.ini", data);
            }

        }

        public void Write(string Object, string Value)
        {
            var parser = new FileIniDataParser();
            var data = parser.ReadFile("Settings.ini");

            data["Settings"][Object] = Value ?? "NA";
            parser.WriteFile("Settings.ini", data);
        }

        private void yPos_TextChanged(object sender, EventArgs e)
        {
            if (!dailyRadio.Checked)
                Write("yPos", yPos.Text);
            else
            {
                var parser = new FileIniDataParser();
                var data = parser.ReadFile("Settings.ini");

                data["Daily_Special"]["yPos"] = yPos.Text ?? "NA";
                parser.WriteFile("Settings.ini", data);
            }
        }

        private void fDelay_TextChanged(object sender, EventArgs e)
        {
            Write("fDelay", fDelay.Text);
        }

        private void jointRadio_CheckedChanged(object sender, EventArgs e)
        {
            UpdateSettings("Joint_Menu");
        }

        private void edibleRadio_CheckedChanged(object sender, EventArgs e)
        {
            UpdateSettings("Edible_Menu");
        }

        private void cartRadio_CheckedChanged(object sender, EventArgs e)
        {
            UpdateSettings("Cart_Menu");
        }

        private void dabRadio_CheckedChanged(object sender, EventArgs e)
        {
            UpdateSettings("Dab_Menu");
        }

        private void UpdateSettings(string Object)
        {
            excelRange1.Text = ReadINI(Object, "Range1").ToUpper();
            excelRange2.Text = ReadINI(Object, "Range2").ToUpper();
            cSizeH.Text = ReadINI(Object, "cH");
            cSizeW.Text = ReadINI(Object, "cW");
            string RadioCheck = ReadINI(Object, "Size");
            if (RadioCheck.ToUpper() == "4K")
            {
                radio_4k.Checked = true;
                radio_1080.Checked = false;
                radio_custom.Checked = false;
            }
            else if (RadioCheck.ToUpper() == "1080")
            {
                radio_4k.Checked = true;
                radio_1080.Checked = false;
                radio_custom.Checked = false;
            }
            else
            {
                radio_4k.Checked = true;
                radio_1080.Checked = false;
                radio_custom.Checked = false;
            }

            if (dailyRadio.Checked)
            {
                xPos.Text = ReadINI("Daily_Special", "xPos");
                yPos.Text = ReadINI("Daily_Special", "yPos");
            }
            else
            {
                xPos.Text = ReadINI("Settings", "xPos");
                yPos.Text = ReadINI("Settings", "yPos");
            }

            autoUpload.Checked = Convert.ToBoolean(ReadINI("Settings", "autoUpload"));
            preview.Checked = Convert.ToBoolean(ReadINI("Settings", "preview"));
        }

        private void SaveSettings(string Object)
        {
            var parser = new FileIniDataParser();
            var data = parser.ReadFile(AppForm.StartupPath + @"\Settings.ini");

            Console.WriteLine("Saving to Settings.ini");
            Console.WriteLine(Object);
            data[Object]["Range1"] = excelRange1.Text.ToUpper() ?? "NA";
            data[Object]["Range2"] = excelRange2.Text.ToUpper() ?? "NA";
            data[Object]["cH"] = cSizeH.Text ?? "NA";
            data[Object]["cW"] = cSizeW.Text ?? "NA";
            if (radio_4k.Checked)
                data[Object]["Size"] = "4K";
            else if (radio_1080.Checked)
                data[Object]["Size"] = "1080";
            else
                data[Object]["Size"] = "Custom";

            parser.WriteFile("Settings.ini", data);
            Console.WriteLine("Saved");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DisposePictureBox();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (!radioButton1.Checked)
                SaveSettings(GetMenuString());
        }

        private const int ColumnBase = 26;
        private const int DigitMax = 7; // ceil(log26(Int32.Max))
        private const string Digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public static string GetLetter(int index)
        {
            if (index <= 0)
                throw new IndexOutOfRangeException("index must be a positive number");

            if (index <= ColumnBase)
                return Digits[index - 1].ToString();

            var sb = new StringBuilder().Append(' ', DigitMax);
            var current = index;
            var offset = DigitMax;
            while (current > 0)
            {
                sb[--offset] = Digits[--current % ColumnBase];
                current /= ColumnBase;
            }
            return sb.ToString(offset, DigitMax - offset);
        }

        public string Merge(int Num_Letter, int Num)
        => (GetLetter(Num_Letter) + Num.ToString());


        public int GetNum(string Range)
        {
            if (Range.Substring(0, 2).ToLower() == "aa")
                return char.ToUpper('z') - 62;

            char c = char.Parse(Range.Substring(0, 1).ToLower());
            return char.ToUpper(c) - 63;
        }

        private List<string> parts = new List<string>();

        private void button10_Click(object sender, EventArgs e)
        {
            parts.Clear();
            if (!radioButton1.Checked)
            {
                try
                {
                    foreach (var process in Process.GetProcessesByName("EXCEL"))
                    {
                        process.Kill();
                        process.WaitForExit();
                    }
                    Console.WriteLine("Turning menu into nice and easy list");
                    string excelFile = AppForm.StartupPath + "\\" + GetMenuString() + ".xlsx";
                    Excel.Application excel = new Excel.Application();
                    Workbook w = excel.Workbooks.Open(excelFile);
                    Worksheet ws = w.Sheets[1];
                    ws.Protect(Contents: false);
                    Console.WriteLine("Menu Loaded");
                    //Console.WriteLine(GetNum(excelRange2.Text));
                    //Console.WriteLine(excelRange2.Text.Substring(1));
                    int c = GetNum(excelRange2.Text);
                    Console.WriteLine(c);
                    int r = Int32.Parse(excelRange2.Text.Substring((c >= 28 ? 2 : 1)));
                    for (int i = 1; i < c; i++)
                    {
                        for (int a = 1; a <= r; a++)
                        {
                            var cellValue = (ws.Cells[a, i] as Range).Value;
                            if (cellValue != null && !FilterString(cellValue))
                                parts.Add(cellValue.ToString());
                        }
                    }

                    w.Close(false);
                    excel.Quit();
                }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                CreateExcel();
            }
            else
                MessageBox.Show("To print the Flower menu press 'Edit' and then Print Menu under the Flower Menu Editor");
        }

        private void CreateExcel()
        {
            Console.WriteLine("Menu scanned creating printable document");
            File.Delete("PrintMenu.xlsx");

            File.Copy("Blank.xlsx", "PrintMenu.xlsx");
            string excelFile = AppForm.StartupPath + "\\PrintMenu.xlsx";
            Excel.Application excel = new Excel.Application();
            Workbook w = excel.Workbooks.Open(excelFile);
            Worksheet ws = w.Sheets[1];
            ws.Protect(Contents: false);
            int i = 0;
            int k = 1;
            foreach (var value in parts)
            {
                i++;
                if (i >= 45)
                {
                    i = 1;
                    k++;
                }
                ws.Cells[i, k].Value = value;
                try
                {
                    int test = Int32.Parse(value.Substring(0, 1));
                    ws.Cells[i, k].Font.Size = 8;
                    ws.Cells[i, k].Font.Bold = true;
                }
                catch { ws.Cells[i, k].Font.Size = 8; }
            }
            w.Save();
            w.Close();
            excel.Quit();
            Console.WriteLine("Printing to default printer");
            SendToPrinter(excelFile);
            //PrintExcel(excelFile);
        }

        private bool FilterString(object str)
        {
            try
            {
                string input = str.ToString().ToLower();
                input = input.Replace(".", "");

                if (input == "a" || input == "cost" || input == "cbd" || input == "thc" || input == "name" || input.Contains("cont") || input == "4oz" || input == "1oz")
                    return true;
                else if (input == "edibles" || input == "joints" || input == "cartridges" || input == "concentrates")
                    return true;
                else
                    return IsDigitsOnly(input);
            }
            catch { return true; }
        }

        public void PrintExcel(string fileName)
        {
            Excel.Application xlexcel = new Excel.Application();
            Workbook xlWorkBook = xlexcel.Workbooks.Open(fileName);
            Worksheet xlWorkSheet = xlWorkBook.Sheets[1];
            Range xlRange = xlWorkSheet.UsedRange;
            object misValue = System.Reflection.Missing.Value;
            // Get the current printer
            string Defprinter = null;
            Defprinter = xlexcel.ActivePrinter;

            // Setup our sheet
            var _with1 = xlWorkSheet.PageSetup;
            // Landscape orientation
            //_with1.Orientation = Excel.XlPageOrientation.xlLandscape;
            // Fit Sheet on One Page 
            _with1.FitToPagesWide = 2;
            _with1.FitToPagesTall = 1;
            // Normal Margins
            _with1.LeftMargin = xlexcel.InchesToPoints(0.7);
            _with1.RightMargin = xlexcel.InchesToPoints(0.7);
            _with1.TopMargin = xlexcel.InchesToPoints(0.75);
            _with1.BottomMargin = xlexcel.InchesToPoints(0.75);
            _with1.HeaderMargin = xlexcel.InchesToPoints(0.3);
            _with1.FooterMargin = xlexcel.InchesToPoints(0.3);

            // Print the range
            xlRange.PrintOutEx(misValue, misValue, misValue, misValue,
            misValue, misValue, misValue, misValue);
        }

        private void SendToPrinter(string File)
        {
            try
            {
                var info = new ProcessStartInfo();
                info.Verb = "print";
                info.FileName = File;
                info.CreateNoWindow = true;
                info.WindowStyle = ProcessWindowStyle.Hidden;

                var p = new Process();
                p.StartInfo = info;
                p.Start();

                p.WaitForInputIdle();
                p.Dispose();
            }
            catch { }
        }

        private bool IsDigitsOnly(object str)
        {
            if (str != null)
                foreach (char c in str.ToString())
                    if (c < '0' || c > '9')
                        return false;
            return true;
        }

        #region Excelless Editor, unused atm. Probably will never use
        private void button11_Click_1(object sender, EventArgs e)
        {
            ExcelessEditor.ExcelessEditor frm = new ExcelessEditor.ExcelessEditor();
            frm.MenuName = GetMenuString();
            frm.Show();
        }
        private char[] alph = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
        private List<string> Companys = new List<string> { };

        private void button12_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("This will delete any existing files being used by the exceless editor! Only do this if it crashes or if the excel file has been updated more recently than the editor!", "Confirm", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (!radioButton1.Checked)
                {
                    try
                    {
                        foreach (var process in Process.GetProcessesByName("EXCEL"))
                        {
                            process.Kill();
                            process.WaitForExit();
                        }

                        Console.WriteLine("Creating files needed for exceless editor....");
                        string excelFile = AppForm.StartupPath + "\\" + GetMenuString() + ".xlsx";
                        string companyFile = AppForm.StartupPath + "\\" + GetMenuString() + "_Companys.txt";
                        string iniFile = AppForm.StartupPath + "\\" + GetMenuString() + "_Products.ini";

                        File.Delete(iniFile);

                        using (FileStream fs = File.Create(iniFile)) { }

                        Companys.Clear();
                        parts.Clear();
                        var parser = new FileIniDataParser();
                        var data = parser.ReadFile(iniFile);

                        Excel.Application excel = new Excel.Application();
                        Workbook w = excel.Workbooks.Open(excelFile);
                        Worksheet ws = w.Sheets[1];
                        ws.Protect(Contents: false);
                        Console.WriteLine("Menu Loaded");
                        int c = GetNum(excelRange2.Text);
                        int r = Int32.Parse(excelRange2.Text.Substring(1));
                        string lastComp = "";
                        for (int i = 1; i < c; i++)
                        {
                            for (int a = 1; a <= r; a++)
                            {
                                var cellValue = (ws.Cells[a, i] as Range).Value;
                                //Oof
                                if (cellValue != null &&
                                    !checkString(cellValue.ToString()) &&
                                    cellValue.ToString() != " " &&
                                    cellValue.ToString() != "" &&
                                    cellValue.ToString() != "Concentrates" &&
                                    cellValue.ToString() != "Cartridges" &&
                                    !IsDigitsOnly(cellValue.ToString()))
                                {//The fuck is this crap
                                    if (cellValue.ToString().Contains(":"))
                                    {
                                        var cellValue1 = (ws.Cells[a, i + 1] as Range).Value;
                                        var cellValue2 = (ws.Cells[a, i + 2] as Range).Value;
                                        string formatted = cellValue.ToString().Replace(": ", ":");
                                        if (cellValue1 != null)
                                            formatted += ":" + cellValue1.ToString();
                                        if (cellValue2 != null)
                                            formatted += ":" + cellValue2.ToString();
                                        parts.Add(formatted);
                                    }
                                    else
                                    {
                                        if (lastComp != "")
                                        {
                                            foreach (string p in parts)
                                                data[lastComp.Split(':')[1]][p.Split(':')[0].ToUpper()] = p;

                                            parser.WriteFile(iniFile, data);
                                        }
                                        Companys.Add(cellValue.ToString().Replace(". ", ":"));
                                        lastComp = cellValue.ToString().Replace(". ", ":");
                                        parts.Clear();
                                    }
                                }
                            }
                        }
                        File.WriteAllText(companyFile, String.Join("\n", Companys));
                        w.Close(false);
                        excel.Quit();
                    }
                    catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                }
                else
                    MessageBox.Show("The Flower menu is already exceless!!");
            }
        }
        #endregion
        private bool checkString(string toCheck)
        {
            foreach (string s in BadWords)
                if (toCheck.ToLower().Contains(s))
                    return true;
            return false;
        }

        private string[] BadWords = { "cost", "thc", "cbd", "name", "cont" };

        private void button13_Click(object sender, EventArgs e)
        {
            Process.Start("https://joexv.github.io/MenuCreatorHelp/");
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            UpdateSettings("Daily_Special");
            button14.Enabled = dailyRadio.Checked;
            if (dailyRadio.Checked)
            {
                groupBox2.Enabled = true;
                fDelay.Enabled = false;
                noExtract.Enabled = false;
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            Process.Start("Daily_Template.png");
        }

        private bool Ping(string IP)
        {
            bool pingable = false;
            Ping pinger = null;
            try
            {
                pinger = new Ping();
                PingReply reply = pinger.Send(IP);
                pingable = reply.Status == IPStatus.Success;
            }
            catch (PingException ex) { }
            finally { if (pinger != null) { pinger.Dispose(); } }

            if (!pingable)
                badEgg = true;

            return pingable;
        }
        private const String APP_ID = "MenuCreator";
        private void OnElapsed(object sender, ElapsedEventArgs e)
        {
            PingAll();
        }

        public void StartTimer(int Minutes = 5)
        {
            timer = new System.Timers.Timer(60000 * Minutes);
            timer.Elapsed += new ElapsedEventHandler(OnElapsed);
            timer.AutoReset = true;
            timer.Start();
        }

        bool badEgg = false;
        private void PingAll()
        {
            badEgg = false;

            sF.Checked = Ping(Flower);
            sJ.Checked = Ping(Joint);
            sE.Checked = Ping(Edible);
            sD.Checked = Ping(Dab);
            sC.Checked = Ping(Cart);
            sS.Checked = Ping(Special);

            if (badEgg)
                Noti("Error!", "One or more of the menus could not be pinged!");
        }

        private void Noti(string Title, string Message)
        {
            notifyIcon1.BalloonTipTitle = Title;
            notifyIcon1.BalloonTipText = Message;
            notifyIcon1.Visible = true;
            notifyIcon1.ShowBalloonTip(180000);
            Thread.Sleep(2000);
            notifyIcon1.Visible = false;
        }

        const string Flower = "192.168.1.112";
        const string Joint = "192.168.1.111";
        const string Edible = "192.168.1.114";
        const string Dab = "192.168.1.115";
        const string Cart = "192.168.1.116";
        const string Special = "192.168.1.67";

        private System.Timers.Timer timer = new System.Timers.Timer();

        private void button15_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Attempting to play the video");
            try
            {
                var task = Task.Run(() => {
                    SSH(@"sudo -u pi nohup omxplayer /home/pi/screenly_assets/movie.mp4 --loop >/dev/null 2>&1 & echo done", "192.168.1.112");
                });

                if (!task.Wait(TimeSpan.FromMilliseconds(10000)))
                    MessageBox.Show("SSH command timed out at 10 seconds. If video is not playing try uploading again and replaying. Make sure video is over 3 seconds long otherwise it will not play!!!");
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {
            CreateMenu();
            DisposePictureBox();
            CreateDailyMenu();
            //Upload(@"C:\Users\Desktop\Joe\Test.png", "192.168.1.67");
            Console.WriteLine("Deleting all old assets");
            DeleteOldAssetsAsync(GetIP());
            Console.WriteLine("Uploading to Screenly Menus");
            DisposePictureBox();

            string Output = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) +
                "\\MenuImages\\" + GetMenuString() + "_" + DateTime.Today.ToString("MM-dd-yyyy") + "_.png";
            SFTPUpload(Output, @"/home/pi/screenly_assets/AUTOMATED", "192.168.1.67");
            Upload(@"/home/pi/screenly_assets/AUTOMATED", "192.168.1.67");
        }

        private Asset AssetToUpdate { get; set; }
        private async void Upload(string fileLocation, string IP)
        {
            Device newDevice = new Device
            {
                Name = "Specials",
                Location = "Floor",
                IpAddress = IP,
                Port = "80",
                ApiVersion = "v1.1/"
            };

            Asset a = new Asset();
            a.AssetId = "AUTOMATED_" + DateTime.Now.ToString("MM-dd-yyyy_hhmm");
            a.Name = "Menu_AUTO_" + DateTime.Now.ToString("MM-dd-yyyy hh-mm tt");
            a.Uri = fileLocation;
            a.StartDate = DateTime.Today.AddDays(-1).ToUniversalTime();
            a.EndDate = DateTime.Today.AddDays(20).ToUniversalTime();
            a.Duration = "10";
            a.IsEnabled = 1;
            a.NoCache = 0;
            a.Mimetype = "image";
            a.SkipAssetCheck = 1;
            a.IsProcessing = 0;
            await newDevice.CreateAsset(a);
        }

        private void autoUpload_CheckedChanged(object sender, EventArgs e)
        {
            Write("autoUpload", autoUpload.Checked.ToString());
        }

        private void preview_CheckedChanged(object sender, EventArgs e)
        {
            Write("preview", preview.Checked.ToString());
        }

        private void button11_Click_2(object sender, EventArgs e)
        {
            button8_Click(sender, e);
            button7_Click(sender, e);
            button5_Click(sender, e);
            button15_Click(sender, e);
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            //Do Commands base on arguments
            if (shouldChange)
            {
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                {
                    process.Kill();
                    process.WaitForExit();
                }
                //MessageBox.Show(Args[0]);
                foreach (string line in Args[0].Replace("&S", " ").Split(new string[] { "&N" }, StringSplitOptions.None))
                {
                    if (line.ToLower().Contains("add")) {
                        addProduct(line);
                    } else if (line.ToLower().Contains("remove")) {
                        removeProduct(line);
                    } else if (line.ToLower().Contains("create")) {
                        createMenu(line);
                    }
                }
                this.Close();
            }
        }

        private static string oauth => File.ReadAllText(@"Z:\Slack Bot\SlackBot_Auth.txt");
        private void print(int[] arr)
        {
            foreach(int i in arr)
                Console.WriteLine(i.ToString());
        }

        private void removeProduct(string cmdLine) { }

        private void createMenu(string cmdLine) { }

        //Unfinished
        private void addProduct(string cmdLine)
        {
            //Prep product specifics
            string[] args = cmdLine.Split(new string[] { " -" }, StringSplitOptions.None);
            string Menu = "";
            Product pdt = new Product();
            foreach (string cmd in args)
            {
                switch (cmd.Substring(0, 2))
                {
                    case "m ":
                        Menu = cmd.Substring(2);
                        break;
                    case "c ":
                        pdt.Company = cmd.Substring(2);
                        break;
                    case "s ":
                        pdt.Name = cmd.Substring(2);
                        break;
                    case "p ":
                        pdt.Cost = cmd.Substring(2);
                        break;
                    case "l ":
                        pdt.Letter = cmd.Substring(2);
                        break;
                    case "th ":
                        pdt.THC = cmd.Substring(3);
                        break;
                    case "cb ":
                        pdt.CBD = cmd.Substring(3);
                        break;
                    case "t ":
                        pdt.Type = cmd.Substring(3);
                        break;
                    default:
                        break;
                }
            }
            Menu = autoMenu(Menu);
            Console.WriteLine(Menu);
            Console.WriteLine(read(pdt));
            int[] settings = getSettings(Menu);
            print(settings);
            //Prep excel file
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wkb = excel.Workbooks.Open(AppForm.StartupPath + "\\" + Menu + ".xlsx");
            Excel.Worksheet sheet = wkb.Worksheets[1] as Excel.Worksheet;
            Range rng = sheet.UsedRange;
            //Row:Column
            Range row = sheet.Rows.Cells[1, settings[1]];
            Range CompanyStart = sheet.Rows.Cells[1, 1];
            Range CompanyEnd = sheet.Rows.Cells[1, 1];
            //Put into try catch to safely close excel for any errors
            try
            {
                bool shouldAdjust = false;
                int noteC = 0;
                Console.WriteLine("Excel file prepped, starting search");
                //Scans for the needed Company
                for (int i = settings[0]; i <= settings[3]; i++)
                {
                    row = sheet.Rows.Cells[i, settings[1]];
                    //If its a company and contains the correct company
                    try
                    {
                        if (row.Value != null || row.Value != "")
                        {
                            Console.WriteLine(row.Value);
                            if (row.Value.Contains(pdt.Company))
                            {
                                CompanyStart = row;
                                noteC = i;
                                //stops for loop in a shitty way
                                i = settings[3] + 2;
                            }
                        }
                    }
                    catch { }

                    if (i == settings[3] && noteC == 0)
                    {
                        //If hits the last line, adjusts to the next row of products
                        Console.WriteLine("Next Row");
                        settings[1] = settings[1] + 1;
                        i = settings[0];
                    }

                    //Reached end of menu
                    if(settings[1] > 23) { Console.Write("End of the menu reached. You need to add this product in manually."); }
                }
                Console.WriteLine("Company found, looking for available letter");
                List<Char> usedLetters = new List<Char>();
                bool doesntExist = true;
                //Scans company for all used letters, and checks to see if the next company needs to be moved to add another product.
                for (int i = noteC + 1; i <= settings[3]; i++)
                {
                    Console.WriteLine(String.Format("{0}:{1}", i, settings[1]));
                    row = sheet.Rows.Cells[i, settings[1]];
                    string value = row.Value;
                    //If the value is empty or null it marks as the end of the company and marks that no changes need to be made to the file as a whole
                    if (row.Value == "" || row.Value == null && row.Interior.Color != 293142)
                    { CompanyEnd = row; shouldAdjust = false; i = settings[3] + 1; }
                    else if (value.Contains("."))
                    {
                        //Another company
                        shouldAdjust = true;
                        CompanyEnd = row;
                        i = settings[3] + 1;
                    }
                    else if (value.Contains(":") || value.Contains(";"))
                    {
                        //Get letter
                        usedLetters.Add(value.ToCharArray()[0]);
                        Console.WriteLine(value.ToCharArray()[0]);
                        if (value.ToLower().Contains(pdt.Name.ToLower()))
                            doesntExist = false;
                    }
                    //Checks if it reaches the end of the menu, by checking for the dark color I use as a border
                    else if (row.Interior.Color == 293142) { shouldAdjust = true; }
                }
                char gLetter = 'Z';
                //Insert in first blank space without adjustment
                if (!shouldAdjust && doesntExist)
                {
                    int i = 0;
                    //Get free letter & make note of how many it takes to find it in order ot inject it and adjust the rest of the products
                    foreach (char c in Digits.ToCharArray())
                    {
                        if (!usedLetters.Contains(c))
                        {
                            gLetter = c;
                            break;
                        }
                        else
                            i++;
                    }
                    Console.WriteLine("Letter found, writing to cell");
                    CompanyEnd.Value = String.Format("    {0}: {1}", gLetter.ToString().ToUpper(), pdt.Name);
                    CompanyEnd.Font.Size = 24;
                    CompanyEnd.Font.Bold = true;
                    CompanyEnd.Font.Color = Color.DarkGreen;
                    CompanyEnd.VerticalAlignment = XlVAlign.xlVAlignCenter;


                    CompanyEnd = sheet.Rows.Cells[CompanyEnd.Row, CompanyEnd.Column + settings[2]];
                    CompanyEnd.Value = pdt.Cost;
                    CompanyEnd.Font.Size = 24;
                    CompanyEnd.Font.Bold = true;
                    CompanyEnd.Font.Color = Color.DarkGreen;
                    CompanyEnd.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    CompanyEnd.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    CompanyEnd.NumberFormat = "$##.00";
                    var client = new SlackClient(oauth);
                    client.PostMessage(string.Format("Product was added! {0}:{1}", pdt.Company, pdt.Name), channel: "#menu_updates", username: "menubot");
                }
                else
                {
                    var client = new SlackClient(oauth);
                    client.PostMessage(string.Format("Prodcut failed to get added! It either exists already, or it requires functions that arent implemented yet! {0}:{1}", pdt.Company, pdt.Name), channel: "#menu_updates", username: "menubot");
                }
            }
            catch (Exception ex)
            {
                var client = new SlackClient(oauth);
                client.PostMessage(string.Format("Prodcut failed to get added! Please attempt to add the product manually! {0}:{1}", pdt.Company, pdt.Name), channel: "#menu_updates", username: "menubot");
            }
            wkb.Save();
            wkb.Close();
            excel.Quit();
        }

        private int[] getSettings(string Menu)
        {
            switch (Menu)
            {
                case "Dab_Menu":
                    return dabSettings;
                case "Cart_Menu":
                    return cartSettings;
                case "Joint_Menu":
                    return jointSettings;
                case "Edible_Menu":
                    return edibleSettings;
                default:
                    return dabSettings;
            }
        }

        //Start Offset(Column)
        //Start Row
        //Space between Name and cost
        //Height of Row -Including Start Offset
        //Space between Name and next row of names
        private int[] dabSettings = { 5, 3, 2, 53, 3};
        private int[] cartSettings = { 5, 3, 2, 53, 3 };
        private int[] jointSettings = { 5, 3, 2, 53, 3 };
        private int[] edibleSettings = { 5, 3, 2, 53, 3 };

        private string autoMenu(string cmd)
        {
            Console.WriteLine(cmd);
            switch (cmd.ToLower())
            {
                case "dabs":
                    return "Dab_menu";
                case "carts":
                    return "Cart_Menu";
                case "edibles":
                    return "Edible_Menu";
                case "joints":
                    return "Joint_Menu";
                case "dab":
                    return "Dab_menu";
                case "cart":
                    return "Cart_Menu";
                case "edible":
                    return "Edible_Menu";
                case "joint":
                    return "Joint_Menu";
                case "d":
                    return "Dab_menu";
                case "c":
                    return "Cart_Menu";
                case "e":
                    return "Edible_Menu";
                case "j":
                    return "Joint_Menu";
                case "test":
                    return "Test_Dab_Menu";
                default:
                    Console.WriteLine(cmd.ToLower());
                    return "Test_Dab_Menu";
            }
        }

        private struct Product
        {
            public string Name;
            public string Cost;
            public string Company;
            public string Letter;
            public string THC;
            public string CBD;
            public string Type;
        }

        private string read(Product pdt)
            => String.Format("{0} {1} {2} {3} {4} {5} {6}", pdt.Name, pdt.Cost, pdt.Company, pdt.Letter, pdt.THC, pdt.CBD, pdt.Type);
    }

    public class Asset
    {
        [Newtonsoft.Json.JsonProperty(PropertyName = "asset_id")]
        public string AssetId { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "mimetype")]
        public string Mimetype { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "name")]
        public string Name { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "end_date")]
        public DateTime EndDate { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "is_enabled")]
        public Int32 IsEnabled { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "is_processing")]
        public Int32? IsProcessing { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "skip_asset_check")]
        public Int32 SkipAssetCheck { get; set; }

        [Newtonsoft.Json.JsonIgnore]
        public bool IsEnabledSwitch
        {
            get
            {
                return IsEnabled.Equals(1) ? true : false;
            }
        }

        [Newtonsoft.Json.JsonProperty(PropertyName = "nocache")]
        public Int32 NoCache { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "is_active")]
        public Int32 IsActive { get; set; }

        private string _Uri;

        [Newtonsoft.Json.JsonProperty(PropertyName = "uri")]
        public string Uri
        {
            get { return _Uri; }
            set { _Uri = System.Net.WebUtility.UrlEncode(value); }
        }

        [Newtonsoft.Json.JsonIgnore]
        public string ReadableUri
        {
            get
            {
                return System.Net.WebUtility.UrlDecode(this.Uri);
            }
        }

        [Newtonsoft.Json.JsonProperty(PropertyName = "duration")]
        public string Duration { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "play_order")]
        public Int32 PlayOrder { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "start_date")]
        public DateTime StartDate { get; set; }
    }
    public class Device
    {
        [Newtonsoft.Json.JsonIgnore]
        private List<Asset> Assets;

        [Newtonsoft.Json.JsonIgnore]
        public bool IsUp { get; set; }

        [Newtonsoft.Json.JsonIgnore]
        public ObservableCollection<Asset> ActiveAssets
        {
            get
            {
                return new ObservableCollection<Asset>(this.Assets.FindAll(x => x.IsActive.Equals(1)));
            }
        }

        [Newtonsoft.Json.JsonIgnore]
        public ObservableCollection<Asset> InactiveAssets
        {
            get
            {
                return new ObservableCollection<Asset>(this.Assets.FindAll(x => x.IsActive.Equals(0)));
            }
        }

        [Newtonsoft.Json.JsonProperty(PropertyName = "name")]
        public string Name { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "ip_address")]
        public string IpAddress { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "port")]
        public string Port { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "location")]
        public string Location { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "api_version")]
        public string ApiVersion { get; set; }

        [Newtonsoft.Json.JsonIgnore]
        public string HttpLink
        {
            get
            {
                return $"http://{IpAddress}:{Port}";
            }
        }

        public Device()
        {
            this.Assets = new List<Asset>();
            this.IsUp = false;
        }

        public async Task<bool> IsReachable()
        {
            try
            {
                HttpClient client = new HttpClient();
                client.Timeout = new TimeSpan(0, 0, 1);

                HttpResponseMessage response = await client.GetAsync(this.HttpLink);
                if (response == null || !response.IsSuccessStatusCode)
                {
                    this.IsUp = false;
                    return false;
                }
                else
                {
                    this.IsUp = true;
                    return true;
                }
            }
            catch
            {
                this.IsUp = false;
                return false;
            }
        }


#region Screenly's API methods

        /// <summary>
        /// Get assets trought Screenly API
        /// </summary>
        /// <returns></returns>
        public async Task GetAssetsAsync()
        {
            List<Asset> returnedAssets = new List<Asset>();
            string resultJson = string.Empty;
            string parameters = $"/api/{this.ApiVersion}assets";

            try
            {
                HttpClient request = new HttpClient();
                using (HttpResponseMessage response = await request.GetAsync(this.HttpLink + parameters))
                {
                    resultJson = await response.Content.ReadAsStringAsync();
                }

                if (!resultJson.Equals(string.Empty))
                    this.Assets = JsonConvert.DeserializeObject<List<Asset>>(resultJson);
            }
            catch (Exception ex)
            {
                throw new Exception("Error while getting assets.", ex);
            }
        }

        /// <summary>
        /// Remove specific asset for selected device
        /// </summary>
        /// <param name="assetId">Asset ID</param>
        /// <returns>Boolean for result of execution</returns>
        public async Task<bool> RemoveAssetAsync(string assetId)
        {
            string resultJson = string.Empty;
            string parameters = $"/api/{this.ApiVersion}assets/{assetId}";

            try
            {
                HttpClient request = new HttpClient();
                using (HttpResponseMessage response = await request.DeleteAsync(this.HttpLink + parameters))
                {
                    resultJson = await response.Content.ReadAsStringAsync();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error when asset deleting.", ex);
            }

            return true;
        }

        /// <summary>
        /// Update specific asset
        /// </summary>
        /// <param name="a">Asset to update</param>
        /// <returns>Asset updated</returns>
        public async Task<Asset> UpdateAssetAsync(Asset a)
        {
            Asset returnedAsset = new Asset();
            JsonSerializerSettings settings = new JsonSerializerSettings();
            IsoDateTimeConverter dateConverter = new IsoDateTimeConverter
            {
                DateTimeFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss.fff'Z'"
            };
            settings.Converters.Add(dateConverter);

            string json = JsonConvert.SerializeObject(a, settings);
            var postData = $"model={json}";
            var data = System.Text.Encoding.UTF8.GetBytes(postData);

            string resultJson = string.Empty;
            string parameters = $"/api/{this.ApiVersion}assets/{a.AssetId}";

            try
            {
                HttpClient client = new HttpClient();
                HttpContent content = new ByteArrayContent(data, 0, data.Length);
                content.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");
                using (HttpResponseMessage response = await client.PutAsync(this.HttpLink + parameters, content))
                {
                    resultJson = await response.Content.ReadAsStringAsync();
                }

                if (!resultJson.Equals(string.Empty))
                {
                    returnedAsset = JsonConvert.DeserializeObject<Asset>(resultJson, settings);
                }
            }
            catch (WebException ex)
            {
                using (var stream = ex.Response.GetResponseStream())
                using (var reader = new StreamReader(stream))
                {
                    throw new Exception(reader.ReadToEnd(), ex);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error while updating asset.", ex);
            }

            return returnedAsset;
        }

        /// <summary>
        /// Update order of active assets throught API
        /// </summary>
        /// <param name="newOrder"></param>
        /// <returns></returns>
        public async Task UpdateOrderAssetsAsync(string newOrder)
        {
            var postData = $"ids={newOrder}";
            var data = System.Text.Encoding.UTF8.GetBytes(postData);

            string resultJson = string.Empty;
            string parameters = $"/api/{this.ApiVersion}assets/order";

            try
            {
                HttpClient client = new HttpClient();
                HttpContent content = new ByteArrayContent(data, 0, data.Length);
                content.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");
                using (HttpResponseMessage response = await client.PostAsync(this.HttpLink + parameters, content))
                {
                    resultJson = await response.Content.ReadAsStringAsync();
                }
            }
            catch (WebException ex)
            {
                using (var stream = ex.Response.GetResponseStream())
                using (var reader = new StreamReader(stream))
                {
                    throw new Exception(reader.ReadToEnd(), ex);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error while updating assets order.", ex);
            }
        }

        /// <summary>
        /// Create new asset on Raspberry using API
        /// </summary>
        /// <param name="a">New asset to create on Raspberry</param>
        /// <returns></returns>
        public async Task CreateAsset(Asset a)
        {
            Asset returnedAsset = new Asset();
            JsonSerializerSettings settings = new JsonSerializerSettings();
            IsoDateTimeConverter dateConverter = new IsoDateTimeConverter
            {
                DateTimeFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss.fff'Z'"
            };
            settings.Converters.Add(dateConverter);

            string json = JsonConvert.SerializeObject(a, settings);
            var postData = $"model={json}";
            var data = System.Text.Encoding.UTF8.GetBytes(postData);

            string resultJson = string.Empty;
            string parameters = $"/api/{this.ApiVersion}assets";

            try
            {
                HttpClient client = new HttpClient();
                HttpContent content = new ByteArrayContent(data, 0, data.Length);
                content.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");

                using (HttpResponseMessage response = await client.PostAsync(this.HttpLink + parameters, content))
                {
                    resultJson = await response.Content.ReadAsStringAsync();
                }

                if (!resultJson.Equals(string.Empty))
                    returnedAsset = JsonConvert.DeserializeObject<Asset>(resultJson, settings);
            }
            catch (WebException ex)
            {
                using (var stream = ex.Response.GetResponseStream())
                using (var reader = new StreamReader(stream))
                {
                    throw new Exception(reader.ReadToEnd(), ex);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error while creating asset.", ex);
            }
        }

        /// <summary>
        /// Return asset identified by asset ID in param API
        /// </summary>
        /// <param name="assetId">Asset ID to find on device</param>
        /// <returns></returns>
        public async Task<Asset> GetAssetAsync(string assetId)
        {
            Asset returnedAsset = new Asset();
            JsonSerializerSettings settings = new JsonSerializerSettings();
            IsoDateTimeConverter dateConverter = new IsoDateTimeConverter
            {
                DateTimeFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss.fff'Z'"
            };
            settings.Converters.Add(dateConverter);

            string resultJson = string.Empty;
            string parameters = $"/api/{this.ApiVersion}assets/{assetId}";

            try
            {
                HttpClient request = new HttpClient();
                using (HttpResponseMessage response = await request.GetAsync(this.HttpLink + parameters))
                {
                    resultJson = await response.Content.ReadAsStringAsync();
                }

                if (!resultJson.Equals(string.Empty))
                    return JsonConvert.DeserializeObject<Asset>(resultJson);
            }
            catch (Exception ex)
            {
                throw new Exception("Error while getting assets.", ex);
            }
            return null;
        }

#endregion
    }
}