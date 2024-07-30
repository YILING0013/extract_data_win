using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.Json;
using System.Windows.Forms;
using ImageProcessor.Imaging.Formats;
using ImageProcessor;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;

namespace ImageMetadataExtractor
{
    public partial class MainForm : Form
    {
        private ContextMenuStrip contextMenu;
        private ToolStripMenuItem topMostMenuItem;
        private ToolStripMenuItem aboutMeMenuItem;

        public MainForm()
        {
            InitializeComponent();
            InitializeContextMenu();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void InitializeContextMenu()
        {
            // 创建上下文菜单
            contextMenu = new ContextMenuStrip();
            topMostMenuItem = new ToolStripMenuItem("取消置顶");
            topMostMenuItem.Click += new EventHandler(ToggleTopMost);
            contextMenu.Items.Add(topMostMenuItem);

            aboutMeMenuItem = new ToolStripMenuItem("About Me");
            aboutMeMenuItem.Click += new EventHandler(ShowAboutMe);
            contextMenu.Items.Add(aboutMeMenuItem);

            // 将上下文菜单关联到窗体
            this.ContextMenuStrip = contextMenu;
        }

        private void ToggleTopMost(object sender, EventArgs e)
        {
            // 切换窗体的置顶状态
            this.TopMost = !this.TopMost;
            topMostMenuItem.Text = this.TopMost ? "取消置顶" : "设置置顶";
        }

        private void ShowAboutMe(object sender, EventArgs e)
        {
            Form aboutForm = new Form
            {
                Text = "About Me",
                Size = new Size(400, 300),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowIcon = false
            };

            Label programInfo = new Label
            {
                Text = "程序介绍：\n这是一个图像元数据提取和处理工具。",
                AutoSize = true,
                Location = new Point(10, 10)
            };

            LinkLabel githubLink = new LinkLabel
            {
                Text = "GitHub",
                AutoSize = true
            };
            githubLink.LinkClicked += (s, args) => Process.Start(new ProcessStartInfo
            {
                FileName = "https://github.com/YILING0013/",
                UseShellExecute = true
            });

            LinkLabel blogLink = new LinkLabel
            {
                Text = "afdian",
                AutoSize = true
            };
            blogLink.LinkClicked += (s, args) => Process.Start(new ProcessStartInfo
            {
                FileName = "https://afdian.com/a/lingyunfei",
                UseShellExecute = true
            });

            TableLayoutPanel panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                AutoSize = true
            };

            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));

            panel.Controls.Add(programInfo, 0, 0);
            panel.Controls.Add(githubLink, 0, 1);
            panel.Controls.Add(blogLink, 0, 2);

            githubLink.Anchor = AnchorStyles.None;
            blogLink.Anchor = AnchorStyles.None;

            aboutForm.Controls.Add(panel);
            aboutForm.ShowDialog(this);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form_DragEnter);
            this.DragDrop += new DragEventHandler(Form_DragDrop);
            // 初始设置为置顶
            this.TopMost = true;
        }

        private void Form_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) || e.Data.GetDataPresent(DataFormats.Html))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void Form_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    ProcessFile(file);
                }
            }
            else if (e.Data.GetDataPresent(DataFormats.Html))
            {
                string htmlData = (string)e.Data.GetData(DataFormats.Html);
                string urlOrBase64 = ExtractImageUrlOrBase64FromHtml(htmlData);
                if (!string.IsNullOrEmpty(urlOrBase64))
                {
                    string tempFilePath = DownloadOrDecodeImage(urlOrBase64);
                    if (!string.IsNullOrEmpty(tempFilePath))
                    {
                        ProcessFile(tempFilePath);
                    }
                }
            }
        }

        private string ExtractImageUrlOrBase64FromHtml(string htmlData)
        {
            // 简单解析HTML数据以提取图像URL或Base64编码
            const string imgTagStart = "<img";
            const string srcAttrStart = "src=\"";
            int imgTagIndex = htmlData.IndexOf(imgTagStart, StringComparison.OrdinalIgnoreCase);
            if (imgTagIndex >= 0)
            {
                int srcAttrIndex = htmlData.IndexOf(srcAttrStart, imgTagIndex, StringComparison.OrdinalIgnoreCase);
                if (srcAttrIndex >= 0)
                {
                    int srcAttrValueStart = srcAttrIndex + srcAttrStart.Length;
                    int srcAttrValueEnd = htmlData.IndexOf("\"", srcAttrValueStart, StringComparison.OrdinalIgnoreCase);
                    if (srcAttrValueEnd > srcAttrValueStart)
                    {
                        return htmlData.Substring(srcAttrValueStart, srcAttrValueEnd - srcAttrValueStart);
                    }
                }
            }
            return null;
        }

        private string DownloadOrDecodeImage(string urlOrBase64)
        {
            if (urlOrBase64.StartsWith("data:image/", StringComparison.OrdinalIgnoreCase))
            {
                // 处理Base64编码的图像
                string base64Data = urlOrBase64.Substring(urlOrBase64.IndexOf("base64,") + 7);
                byte[] imageBytes = Convert.FromBase64String(base64Data);
                string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
                File.WriteAllBytes(tempFilePath, imageBytes);
                return tempFilePath;
            }
            else
            {
                // 处理URL
                try
                {
                    using (WebClient client = new WebClient())
                    {
                        string tempFilePath = Path.Combine(Path.GetTempPath(), Path.GetFileName(urlOrBase64));
                        client.DownloadFile(urlOrBase64, tempFilePath);
                        return tempFilePath;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"从URL下载图像时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error, 0, MessageBoxOptions.DefaultDesktopOnly);
                    return null;
                }
            }
        }

        private void BtnOpenFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Image Files|*.png;*.jpg;*.jpeg";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ProcessFile(openFileDialog.FileName);
                }
            }
        }

        private void ProcessFile(string filePath)
        {
            string compressedDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "compressed_images");
            if (!Directory.Exists(compressedDir))
            {
                Directory.CreateDirectory(compressedDir);
            }

            string outputExcel = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AI_images_extracted_data.xlsx");

            try
            {
                var commentData = ReadHiddenDataFromImageUsingPython(filePath);
                if (commentData == null)
                {
                    MessageBox.Show("未从图像上读取到隐藏信息，", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error,0,MessageBoxOptions.DefaultDesktopOnly);
                    return;
                }

                string imgName = Path.GetFileNameWithoutExtension(filePath);
                string compressedImgPathJpeg = Path.Combine(compressedDir, imgName + ".jpg");
                string compressedImgPathPng = Path.Combine(compressedDir, imgName + ".png");
                CompressImage(filePath, compressedImgPathJpeg, compressedImgPathPng);

                using (var package = new ExcelPackage(new FileInfo(outputExcel)))
                {
                    var ws = package.Workbook.Worksheets.FirstOrDefault() ?? package.Workbook.Worksheets.Add("Extracted Data");

                    if (ws.Dimension == null)
                    {
                        var headers = new string[]
                        {
                "Image Path", "prompt", "uc", "steps", "height", "width", "scale", "uncond_scale", "cfg_rescale", "seed",
                "n_samples", "hide_debug_overlay", "noise_schedule", "legacy_v3_extend", "reference_information_extracted",
                "reference_strength", "sampler", "controlnet_strength", "controlnet_model", "dynamic_thresholding",
                "dynamic_thresholding_percentile", "dynamic_thresholding_mimic_scale", "sm", "sm_dyn", "skip_cfg_below_sigma",
                "lora_unet_weights", "lora_clip_weights", "request_type", "signed_hash"
                        };

                        for (int i = 0; i < headers.Length; i++)
                        {
                            ws.Cells[1, i + 1].Value = headers[i];
                            ws.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.Gold);
                            ws.Cells[1, i + 1].Style.Font.Bold = true;
                            ws.Cells[1, i + 1].Style.Font.Color.SetColor(Color.White);
                            ws.Cells[1, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            ws.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            ws.Cells[1, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                    }

                    int row = ws.Dimension.End.Row + 1;
                    ws.Cells[row, 1].Value = compressedImgPathJpeg;
                    int col = 2;
                    foreach (var header in ws.Cells[1, 2, 1, ws.Dimension.End.Column].Select(c => c.Text))
                    {
                        ws.Cells[row, col++].Value = commentData.ContainsKey(header) ? commentData[header]?.ToString() : "";
                    }

                    // 添加图像到Excel中
                    string uniqueImgName = imgName + "_" + Guid.NewGuid().ToString();
                    var excelImage = ws.Drawings.AddPicture(uniqueImgName, new FileInfo(compressedImgPathJpeg));
                    int maxImageWidth = 500;  // 设置最大图像宽度
                    int maxImageHeight = 500; // 设置最大图像高度

                    using (var pilImg = System.Drawing.Image.FromFile(compressedImgPathJpeg))
                    {
                        int imgWidth = pilImg.Width;
                        int imgHeight = pilImg.Height;
                        float scale = Math.Min(maxImageWidth / (float)imgWidth, maxImageHeight / (float)imgHeight);
                        excelImage.SetSize((int)(imgWidth * scale), (int)(imgHeight * scale));
                    }

                    excelImage.SetPosition(row - 1, 0, 0, 0);

                    // 设置列颜色和边框
                    string[] columnColors = { "FFEBEE", "FFF3E0", "E8F5E9", "E3F2FD", "F3E5F5", "FFEBEE", "FFF3E0", "E8F5E9", "E3F2FD", "F3E5F5",
                                  "FFEBEE", "FFF3E0", "E8F5E9", "E3F2FD", "F3E5F5", "FFEBEE", "FFF3E0", "E8F5E9", "E3F2FD", "F3E5F5",
                                  "FFEBEE", "FFF3E0", "E8F5E9", "E3F2FD", "F3E5F5", "FFEBEE", "FFF3E0", "E8F5E9", "E3F2FD" };

                    for (int colNum = 1; colNum <= ws.Dimension.End.Column; colNum++)
                    {
                        for (int rowNum = 2; rowNum <= ws.Dimension.End.Row; rowNum++)
                        {
                            var cell = ws.Cells[rowNum, colNum];
                            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml($"#{columnColors[colNum - 1]}"));
                            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            cell.Style.WrapText = true;
                        }
                    }

                    // 调整列宽和行高
                    ws.Column(1).Width = maxImageHeight * 0.15; // 根据行高设置列宽
                    for (int i = 2; i <= ws.Dimension.End.Column; i++)
                    {
                        ws.Column(i).Width = 20;
                    }
                    for (int i = 2; i <= ws.Dimension.End.Row; i++)
                    {
                        ws.Row(i).Height = maxImageHeight * 0.75;
                    }

                    package.Save();
                }

                ShowOutputExcelMessage(outputExcel);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"处理文件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error,0,MessageBoxOptions.DefaultDesktopOnly);
            }
        }

        private void ShowOutputExcelMessage(string outputExcel)
        {
            Form messageForm = new Form
            {
                Text = "信息",
                Size = new Size(400, 200),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowIcon = false
            };

            Label messageLabel = new Label
            {
                Text = "图像为原图，AI信息已保存到：",
                AutoSize = true,
                Location = new Point(10, 10)
            };

            LinkLabel linkLabel = new LinkLabel
            {
                Text = outputExcel,
                AutoSize = false,
                Size = new Size(360, 60),
                Location = new Point(10, 40),
                TextAlign = ContentAlignment.MiddleLeft
            };
            linkLabel.LinkClicked += (s, e) =>
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = outputExcel,
                    UseShellExecute = true
                });
            };

            Button confirmButton = new Button
            {
                Text = "确认",
                DialogResult = DialogResult.OK,
                Location = new Point(150, 120),
                Size = new Size(100, 30)
            };

            messageForm.Controls.Add(messageLabel);
            messageForm.Controls.Add(linkLabel);
            messageForm.Controls.Add(confirmButton);
            messageForm.AcceptButton = confirmButton; // 设置确认按钮为默认按钮

            messageForm.ShowDialog(this);
        }


        private Dictionary<string, object> ReadHiddenDataFromImageUsingPython(string imagePath)
        {
            // 确保 extract_data.exe 位于当前应用程序目录下
            string scriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "extract_data.exe");
            string outputPath = Path.Combine(Path.GetTempPath(), "output.json");

            if (!File.Exists(scriptPath))
            {
                throw new FileNotFoundException("extract_data.exe not found in application directory.");
            }

            var start = new ProcessStartInfo
            {
                FileName = scriptPath,
                Arguments = $"\"{imagePath}\" \"{outputPath}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
            };

            using (var process = Process.Start(start))
            {
                using (var reader = process.StandardOutput)
                {
                    string result = reader.ReadToEnd();
                    Console.WriteLine(result);
                }

                using (var reader = process.StandardError)
                {
                    string result = reader.ReadToEnd();
                    Console.WriteLine(result);
                }

                process.WaitForExit();
            }

            if (File.Exists(outputPath))
            {
                string json = File.ReadAllText(outputPath);
                var commentData = JsonSerializer.Deserialize<Dictionary<string, object>>(json);
                File.Delete(outputPath);
                return commentData;
            }

            return null;
        }

        private void CompressImage(string inputPath, string outputPathJpeg, string outputPathPng)
        {
            using (var imageFactory = new ImageFactory(preserveExifData: true))
            {
                imageFactory.Load(inputPath)
                            .Format(new JpegFormat { Quality = 80 })
                            .Save(outputPathJpeg);

                imageFactory.Load(inputPath)
                            .Format(new PngFormat())
                            .Save(outputPathPng);
            }
        }
    }
}
