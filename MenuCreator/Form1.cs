using IniParser;
using Microsoft.Office.Interop.Excel;
using Renci.SshNet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WMPLib;
using AppForm = System.Windows.Forms.Application;

using Excel = Microsoft.Office.Interop.Excel;

namespace MenuCreator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {
        }

        #region Everything But Uploading to TV

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
                try
                {
                    string Menu = "";
                    if (jointRadio.Checked)
                    {
                        Menu = "Joint_Menu.xlsx";
                    }
                    else if (edibleRadio.Checked)
                    {
                        Menu = "Edible_Menu.xlsx";
                    }
                    else if (cartRadio.Checked)
                    {
                        Menu = "Cart_Menu.xlsx";
                    }
                    else if (dabRadio.Checked)
                    {
                        Menu = "Dab_menu.xlsx";
                    }

                    Process.Start(Menu);
                }
                catch
                {
                }
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
            {
                Menu = "Joint_Menu";
            }
            else if (edibleRadio.Checked)
            {
                Menu = "Edible_Menu";
            }
            else if (cartRadio.Checked)
            {
                Menu = "Cart_Menu";
            }
            else if (dabRadio.Checked)
            {
                Menu = "Dab_Menu";
            }

            return Menu;
        }

        public bool Export_1080p = false;

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                CreateMenu();
            }
            catch
            {
                Console.WriteLine("Menu creation failed, trying again....");
                DisposePictureBox();
                if (File.Exists("Menu_Small.png"))
                {
                    File.Delete("Menu_Small.png");
                }

                FileName = GetMenuString();
                Output = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\" + FileName + "_" + DateTime.Today.ToString("MM-dd-yyyy") + "_.png";
                if (File.Exists(Output))
                {
                    File.Delete(Output);
                }

                CreateMenu();
            }
        }

        public void CreateMenu()
        {
            Cursor.Current = Cursors.WaitCursor;
            foreach (var process in Process.GetProcessesByName("EXCEL"))
            {
                process.Kill();
                process.WaitForExit();
            }

            DisposePictureBox();

            FileName = GetMenuString();
            XLSX = FileName + ".xlsx";
            Output = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\" + FileName + "_" + DateTime.Today.ToString("MM-dd-yyyy") + "_.png";
            FileLocation = System.Windows.Forms.Application.StartupPath + "\\" + XLSX;
            if (File.Exists(Output))
            {
                File.Delete(Output);
            }

            if (CheckForFile(XLSX))
            {
                ConvertFile(Output);
            }

            PreviewImage();
            Cursor.Current = Cursors.Default;
        }

        public void ConvertFile(string output)
        {
            Clipboard.Clear();
            CreateImage_Alt(excelRange1.Text, excelRange2.Text);
            ResizeImage("Menu_Small.png");

            //File.Delete("Menu_Small.png");

            if (File.Exists(output))
            {
                MessageBox.Show("Image Created! It should be located on your desktop.");
            }
            else
            {
                Console.WriteLine("Image creation failed....");
                MessageBox.Show("Looks like something went wrong.");
            }
        }

        public string ReadINI(string Key = "Settings", string Object = "")
        {
            var parser = new FileIniDataParser();
            var data = parser.ReadFile("Settings.ini");
            return data[Key][Object];
        }

        public bool CheckForFile(string ExcelFile)
        {
            bool Temp = File.Exists(ExcelFile);
            if (!Temp)
            {
                MessageBox.Show("Excel file is missing!");
            }

            return Temp;
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
            {
                Console.WriteLine("Disposing picture box image...");
                try
                {
                    var image = pictureBox1.Image;
                    pictureBox1.Image = null;
                    image.Dispose();
                }
                catch
                {
                    Console.WriteLine(
                        "Failed to dispose of picture box image, probably due to picture box not having an image.");
                }
            }
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
                {
                    resizedImage.Save(Output);
                }
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
            {
                w.Write(r.ReadBytes((int)s.Length));
            }
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
            string Picture = "";
            try
            {
                if (jointRadio.Checked)
                {
                    Picture = "Joint_Menu";
                }
                else if (edibleRadio.Checked)
                {
                    Picture = "Edible_Menu";
                }
                else if (cartRadio.Checked)
                {
                    Picture = "Cart_Menu";
                }
                else if (dabRadio.Checked)
                {
                    Picture = "Dab_Menu";
                }
                else
                {
                    if (radioButton1.Checked)
                    {
                        Picture = "Menu";
                    }
                }
                string ImageLocation = Environment.GetFolderPath
                                           (Environment.SpecialFolder.DesktopDirectory) + "\\" + Picture + "_" +
                                       DateTime.Today.ToString("MM-dd-yyyy") + "_.png";
                pictureBox1.Image = Image.FromFile(ImageLocation);
            }
            catch
            {
            }
        }

        #endregion Everything But Uploading to TV

        //Upload to TV
        private void button4_Click(object sender, EventArgs e)
        {
            Process.Start(@"Z:\Menus - OLD\Menu Manager\index.html");
        }

        private void OldWeb()
        {
            try
            {
                string URL = "";
                if (jointRadio.Checked)
                {
                    URL = "192.168.1.111";
                }
                else if (edibleRadio.Checked)
                {
                    URL = "192.168.1.114";
                }
                else if (cartRadio.Checked)
                {
                    URL = "192.168.1.115";
                }
                else if (dabRadio.Checked)
                {
                    URL = "192.168.1.116";
                }

                Process.Start("Http://" + URL);
            }
            catch { }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (!File.Exists("Settings.ini"))
            {
                Extract("MenuCreator", System.Windows.Forms.Application.StartupPath, "Files", "Settings.ini");
            }

            if (File.Exists("Menu_Small.png"))
            {
                File.Delete("Menu_Small.png");
            }

            _writer = new TextBoxStreamWriter(txtConsole);
            Console.SetOut(_writer);

            xPos.Text = ReadINI("Settings", "xPos");
            yPos.Text = ReadINI("Settings", "yPos");
            fDelay.Text = ReadINI("Settings", "fDelay");
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            button2.Enabled = !radioButton1.Checked;
            button5.Enabled = radioButton1.Checked;
            button7.Enabled = radioButton1.Checked;
            button8.Enabled = radioButton1.Checked;
            groupBox3.Enabled = !radioButton1.Checked;
            groupBox2.Enabled = radioButton1.Checked;
        }

        /*
         * Uploads 'movie.mp4' to the Flower Menu Raspberry Pi and plays it.
         * Done entirely in ssh.
         * Timeout was added because the ssh command wasn't returning a completion.
         * Typically uploads a 5MB video in a second or two, but results may differ on your internet
         */

        private void button5_Click(object sender, EventArgs e)
        {
            if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\movie.mp4"))
            {
                Cursor.Current = Cursors.WaitCursor;
                Console.WriteLine("Uploading movie.mp4");
                ScpUpload(Environment.GetFolderPath
                              (Environment.SpecialFolder.DesktopDirectory) + "\\movie.mp4",
                    @"/home/pi/screenly_assets/movie.mp4", "192.168.1.112");
                Console.WriteLine("Movie should be uploaded, runnning ssh command");

                var task = Task.Run(() =>
                {
                    SSH(@"sudo -u pi nohup omxplayer /home/pi/screenly_assets/movie.mp4 --loop >/dev/null 2>&1 & echo done", "192.168.1.112");
                });

                bool isCompletedSuccessfully = task.Wait(TimeSpan.FromMilliseconds(3000));
                if (!isCompletedSuccessfully)
                {
                    MessageBox.Show("Done. Please note that this is experimental and may not function 100%.");
                }

                Cursor.Current = Cursors.Default;
            }
            else
            {
                MessageBox.Show(
                    "There's no movie named 'movie.mp4' on your desktop. If you don't know what you're doing you shouldn't be running this command.");
            }
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
            string IP = "";
            try
            {
                if (jointRadio.Checked)
                {
                    IP = "192.168.1.111";
                }
                else if (edibleRadio.Checked)
                {
                    IP = "192.168.1.114";
                }
                else if (cartRadio.Checked)
                {
                    IP = "192.168.1.115";
                }
                else if (dabRadio.Checked)
                {
                    IP = "192.168.1.116";
                }
                else if (radioButton1.Checked)
                {
                    IP = "192.168.1.112";
                }

                SSH("sudo reboot", IP);
            }
            catch
            {
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var task = Task.Run(() =>
            {
                SSH(@"sudo kill -9 `pgrep omxplayer` & echo done", "192.168.1.112");
            });
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
                {
                    Process.WaitForExit();
                }

                Process.Close();
                Process.Dispose();
                Process = null;
            }
            catch (Exception e)
            {
                Console.WriteLine("Error Processing ExecuteCommand : " + e.Message);
            }
            finally
            {
                //Process = null;
                //ProcessInfo = null;
            }
        }

        private static void Process_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            Console.WriteLine(e.Data);
        }

        private static void SortOutputHandler(object sendingProcess, DataReceivedEventArgs outLine)
        {
            if (!String.IsNullOrEmpty(outLine.Data))
            {
                cmdOutput.Append(Environment.NewLine + outLine.Data);
            }
        }

        private bool BootTooBig()
        {
            string fileLocation = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\Menu_" + DateTime.Today.ToString("MM-dd-yyyy") + "_.png";
            using (Image img = Image.FromFile(fileLocation))
            {
                if (img.Width > 1920 || img.Height > 1080)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            Console.WriteLine("Starting Video creation process. This can take upwards of 10 minutes depending on your computer, GIF size/length and image size.");
            if (!BootTooBig())
            {
                VideoCreation();
            }
            else
            {
                DialogResult dialogResult = MessageBox.Show("Your image is larger than 1080p, do you want to continue making this into a video? Doing so can make the process take considerablly longer or even cause it to fail if your computer cant handle larger resolutions.", "Export options", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    VideoCreation();
                }
                else
                {
                    Console.WriteLine("Video Creation cancelled");
                }
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
            Output = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\Menu_" + DateTime.Today.ToString("MM-dd-yyyy") + "_.png";
            if (File.Exists(Path.Combine(Desktop, Output)) && File.Exists(Path.Combine(Desktop, "Insert.gif")))
            {
                String Command = "";
                if (!noExtract.Checked)
                {
                    try
                    {
                        Directory.Delete(Path.Combine(Desktop, "Gif"), true);
                        Console.WriteLine("Deleted old GIF");
                    }
                    catch
                    {
                    }
                    Console.WriteLine("Both the GIF and menu are located extracting gif");
                    Directory.CreateDirectory(Path.Combine(Desktop, "Gif"));
                    Command = @"cd " + ImageMagick + " & magick convert -coalesce " + Path.Combine(Desktop, "Insert.gif") + " " + Path.Combine(Desktop, "Gif\\Target.png");
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
                Command = @"cd " + ImageMagick + " & magick convert -delay " + fDelay.Text + " -loop 0 " + Path.Combine(Desktop, "Gif\\Target-*.png") + " C:\\Users\\Joe\\Desktop\\Menu_GIF.gif";
                cmd(Command, true, true, true);

                Console.WriteLine(Path.Combine(Desktop, "Output_Gif_" + DateTime.Today.ToString("MM-dd-yyyy")) + ".gif");
                Console.WriteLine("Creating mp4 file out of GIF");
                Command = AppForm.StartupPath + "\\HandBrakeCLI.exe -Z \"Very Fast 1080p30\" -i C:\\Users\\Joe\\Desktop\\Menu_GIF.gif -o " + Path.Combine(Desktop, "movie.mp4") + " & echo done";
                cmd(Command, true, true, true);

                try
                {
                    if (Duration(Path.Combine(Desktop, "movie.mp4")) <= 5)
                    {
                        Console.WriteLine("Video is 5 seconds or under extending...");
                        CombineVideos();
                        Console.WriteLine("Done");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    Console.WriteLine("Error on combining videos!");
                }

                try
                {
                    Directory.Delete(Path.Combine(Desktop, "Gif"), true);
                }
                catch
                {
                    Console.WriteLine("Error on removing leftover files, files might have already been deleted or are still in use.");
                }
            }
            else
            {
                MessageBox.Show("The menu needs to be on your desktop exactly as it was when it was created, and your gif needs to be on your desktop named 'insert.gif'.");
            }
            Cursor.Current = Cursors.Default;
        }

        public void CombineVideos()
        {
            string Desktop = Environment.GetFolderPath
                (Environment.SpecialFolder.DesktopDirectory);

            if (File.Exists(AppForm.StartupPath + "\\movie.mp4"))
                File.Delete(AppForm.StartupPath + "\\movie.mp4");

            if (File.Exists(Path.Combine(AppForm.StartupPath, "movie_combined.mp4")))
                File.Delete(Path.Combine(AppForm.StartupPath, "movie_combined.mp4"));
            

            File.Copy(Path.Combine(Desktop, "movie.mp4"), AppForm.StartupPath + "\\movie.mp4");

            string Commands = "-safe 0 -f concat -i list.txt -c copy " + Path.Combine(AppForm.StartupPath, "movie_combined.mp4");
            Process p = new Process();
            p.StartInfo.FileName = AppForm.StartupPath + "\\ffmpeg.exe";
            p.StartInfo.Arguments = Commands;
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.CreateNoWindow = true;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.Verb = "runas";
            p.Start();
     
            p.WaitForExit();
            p.Close();
            p.Dispose();
            p = null;
            if(File.Exists(Path.Combine(Desktop, "movie.mp4")))
                File.Delete(Path.Combine(Desktop, "movie.mp4"));

            File.Move(Path.Combine(AppForm.StartupPath, "movie_combined.mp4"), Path.Combine(Desktop, "movie.mp4"));
            if (File.Exists("movie.mp4"))
                File.Delete("movie.mp4");

            if (Duration(Path.Combine(Desktop, "movie.mp4")) <= 5)
                CombineVideos();
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
            Write("xPos", xPos.Text);
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
            Write("yPos", yPos.Text);
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
            {
                data[Object]["Size"] = "4K" ?? "NA";
            }
            else if (radio_1080.Checked)
            {
                data[Object]["Size"] = "1080" ?? "NA";
            }
            else
            {
                data[Object]["Size"] = "Custom" ?? "NA";
            }

            parser.WriteFile("Settings.ini", data);
            Console.WriteLine("Saved");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DisposePictureBox();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string Menu = "";
            if (jointRadio.Checked)
            {
                Menu = "Joint_Menu";
            }
            else if (edibleRadio.Checked)
            {
                Menu = "Edible_Menu";
            }
            else if (cartRadio.Checked)
            {
                Menu = "Cart_Menu";
            }
            else if (dabRadio.Checked)
            {
                Menu = "Dab_Menu";
            }
            if (!radioButton1.Checked)
            {
                SaveSettings(Menu);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
        }

        public string GetIP()
        {
            string URL = "";
            if (jointRadio.Checked)
            {
                URL = "192.168.1.111";
            }
            else if (edibleRadio.Checked)
            {
                URL = "192.168.1.114";
            }
            else if (cartRadio.Checked)
            {
                URL = "192.168.1.115";
            }
            else if (dabRadio.Checked)
            {
                URL = "192.168.1.116";
            }
            else if (radioButton1.Checked)
            {
                URL = "192.168.1.112";
            }
            return URL;
        }

        private const int ColumnBase = 26;
        private const int DigitMax = 7; // ceil(log26(Int32.Max))
        private const string Digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public static string GetLetter(int index)
        {
            if (index <= 0)
            {
                throw new IndexOutOfRangeException("index must be a positive number");
            }

            if (index <= ColumnBase)
            {
                return Digits[index - 1].ToString();
            }

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
        {
            return (GetLetter(Num_Letter) + Num.ToString());
        }

        public int GetNum(string Range)
        {
            char c = char.Parse(Range.Substring(0, 1).ToLower());
            return char.ToUpper(c) - 63;
        }

        private List<string> parts = new List<string>();

        private void button10_Click(object sender, EventArgs e)
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
                    int r = Int32.Parse(excelRange2.Text.Substring(1));
                    for (int i = 1; i < c; i++)
                    {
                        for (int a = 1; a <= r; a++)
                        {
                            //Console.WriteLine(a);
                            var cellValue = (ws.Cells[a, i] as Range).Value;
                            if (cellValue != null)
                            {
                                if (!FilterString(cellValue))
                                {
                                    if (cellValue.ToString().ToLower() != "name")
                                    {
                                        parts.Add(cellValue.ToString());
                                    }
                                }
                            }
                        }
                    }

                    w.Close(false);
                    excel.Quit();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                CreateExcel();
            }
            else
            {
                MessageBox.Show("To print the Flower menu press 'Edit' and then Print Menu under the Flower Menu Editor");
            }
        }

        private void CreateExcel()
        {
            Console.WriteLine("Menu scanned creating printable document");
            if (File.Exists("PrintMenu.xlsx"))
            {
                File.Delete("PrintMenu.xlsx");
            }

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
        }

        private bool FilterString(object str)
        {
            try
            {
                string input = str.ToString().ToLower();
                input = input.Replace(".", "");

                if (input == "cost")
                {
                    return true;
                }
                else
                {
                    return IsDigitsOnly(input);
                }
            }
            catch
            {
                return true;
            }
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
                //Thread.Sleep(3000);
                //if (false == p.CloseMainWindow())
                //    p.Kill();
            }
            catch { }
        }

        private bool IsDigitsOnly(object str)
        {
            if (str != null)
            {
                foreach (char c in str.ToString())
                {
                    if (c < '0' || c > '9')
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {
        }

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
                        string excelFile = System.Windows.Forms.Application.StartupPath + "\\" + GetMenuString() + ".xlsx";
                        string companyFile = System.Windows.Forms.Application.StartupPath + "\\" + GetMenuString() + "_Companys.txt";
                        string iniFile = System.Windows.Forms.Application.StartupPath + "\\" + GetMenuString() + "_Products.ini";

                        if (File.Exists(iniFile))
                        {
                            File.Delete(iniFile);
                        }

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
                                if (cellValue != null &&
                                    !checkString(cellValue.ToString()) &&
                                    cellValue.ToString() != " " &&
                                    cellValue.ToString() != "" &&
                                    cellValue.ToString() != "Concentrates" &&
                                    cellValue.ToString() != "Cartridges" &&
                                    !IsDigitsOnly(cellValue.ToString()))
                                {
                                    if (cellValue.ToString().Contains(":"))
                                    {
                                        var cellValue1 = (ws.Cells[a, i + 1] as Range).Value;
                                        var cellValue2 = (ws.Cells[a, i + 2] as Range).Value;
                                        string formatted = cellValue.ToString().Replace(": ", ":");
                                        if (cellValue1 != null)
                                        {
                                            formatted += ":" + cellValue1.ToString();
                                        }

                                        if (cellValue2 != null)
                                        {
                                            formatted += ":" + cellValue2.ToString();
                                        }

                                        parts.Add(formatted);
                                    }
                                    else
                                    {
                                        if (lastComp != "")
                                        {
                                            foreach (string p in parts)
                                            {
                                                data[lastComp.Split(':')[1]][p.Split(':')[0].ToUpper()] = p;
                                            }

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
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }
                    //CreateExcel();
                }
                else
                {
                    MessageBox.Show("The Flower menu is already exceless!!");
                }
            }
        }

        private bool checkString(string toCheck)
        {
            foreach (string s in BadWords)
            {
                if (toCheck.ToLower().Contains(s))
                {
                    return true;
                }
            }

            return false;
        }

        private string[] BadWords = { "cost", "thc", "cbd", "name", "cont" };
    }
}