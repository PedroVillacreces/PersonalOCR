using System;
using System.Collections.Generic;
using System.IO;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using IronOcr;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        public string path { get; set; } =  @"C:\Digitalizacion.txt";
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            InitializeOpenFileDialog();
        }

        private void InitializeOpenFileDialog()
        {
            // Set the file dialog to filter for graphics files.
            this.openFileDialog1.Filter =
                "Images (*.BMP;*.JPG;*.GIF;*.PNG)|*.BMP;*.JPG;*.GIF;*.PNG|" +
                "All files (*.*)|*.*";

            //  Allow the user to select multiple images.
            this.openFileDialog1.Multiselect = true;
            //                   ^  ^  ^  ^  ^  ^  ^

            this.openFileDialog1.Title = "My Image Browser";
        }

        private void Calculate(int i)
        {
            var pow = Math.Pow(i, i);
        }

        public void DoWork(IProgress<int> progress)
        {
            for (var j = 0; j < 100000; j++)
            {
                Calculate(j);

                progress?.Report((j + 1) * 100 / 100000);
            }
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            var dr = this.openFileDialog1.ShowDialog();
            if (dr != DialogResult.OK) return;
            var app = new Microsoft.Office.Interop.Word.Application();
            this.CreateDocument();
            var doc =
                app.Documents.Open(path);
            object missing = System.Reflection.Missing.Value;
            var filePath = new List<string>();
            foreach (var file in openFileDialog1.FileNames)
            {
                try
                {
                   filePath.Add(file);
                }
                catch (SecurityException ex)
                {
                    MessageBox.Show("Security error. Please contact your administrator for details.\n\n" +
                                    "Error message: " + ex.Message + "\n\n" +
                                    "Details (send to Support):\n\n" + ex.StackTrace
                    );
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Cannot display the image: " + file.Substring(file.LastIndexOf('\\'))
                                                                 + ". You may not have permission to read the file, or " +
                                                                 "it may be corrupt.\n\nReported error: " + ex.Message);

                }
            }
            progressBar1.Maximum = 100;
            progressBar1.Step = 1;

            var progress = new Progress<int>(v =>
            {
                progressBar1.Value = v;
            });

            await Task.Run(() => DoWork(progress));
            var ocr = new AdvancedOcr()
            {
                Language = IronOcr.Languages.Spanish.OcrLanguagePack,
                ColorSpace = AdvancedOcr.OcrColorSpace.GrayScale,
                EnhanceResolution = true,
                EnhanceContrast = true,
                CleanBackgroundNoise = true,
                ColorDepth = 4,
                RotateAndStraighten = false,
                DetectWhiteTextOnDarkBackgrounds = false,
                ReadBarCodes = false,
                Strategy = AdvancedOcr.OcrStrategy.Fast,
                InputImageType = AdvancedOcr.InputTypes.Document
            };
            var result = ocr.Read(filePath);
            doc.Content.Text = result.Text;
            doc.Save();
            doc.Close(ref missing);
            app.Quit(ref missing);
            MessageBox.Show("Archivos procesados correctamente");
        }

        private void CreateDocument()
        {
            try
            {
                // Check if file already exists. If yes, delete it.     
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                // Create a new file     
                using (var fs = File.Create(path))
                {
                    // Add some text to file    
                    var title = new UTF8Encoding(true).GetBytes("New Text File");
                    fs.Write(title, 0, title.Length);
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.ToString());
            }
        }
    }
}
