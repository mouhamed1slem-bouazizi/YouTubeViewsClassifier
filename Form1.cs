using System;
using System.Windows.Forms;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Drawing;
using System.Linq;
using System.Collections.Generic;
using ExcelDataReader;
using System.Data;
using System.IO;
using System.Text;
using OfficeOpenXml;
using System.IO;

namespace YouTubeViewsClassifier
{
    
    public partial class Form1 : Form
    {
        static Form1()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
        private DataGridView dataGridView;
        private Button refreshButton;
        private const string API_KEY = "AIzaSyB2MFgLVS1r1rzMMniCWmdwcqOl9B3K9II";
        private readonly HttpClient httpClient;
        private Button importButton;
        private List<string> videoUrls = new List<string>();
        private Button exportButton;
        private Button startCountButton;
        private Label timerLabel;
        private System.Windows.Forms.Timer countdownTimer;
        private int currentDay = 0;
        private int hoursRemaining = 168; // 7 days * 24 hours
        private DateTime lastRefreshTime;
        private int minutesRemaining = 168 * 60; // Convert hours to minutes (7 days * 24 hours * 60 minutes)

        public Form1()
        {
            InitializeComponent();
            httpClient = new HttpClient();
            SetupUI();
            // Load data after UI is setup
            this.Shown += async (s, e) => await Task.Run(async () => await LoadSampleData());
        }

        private void SetupUI()
        {
            this.Size = new Size(1200, 600);
            this.Text = "YouTube Views Classifier";

            // Create main table layout
            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(10)
            };

            // Setup top controls panel
            Panel topPanel = new Panel
            {
                Height = 40,
                Dock = DockStyle.Fill
            };

            // Create and setup Refresh button
            refreshButton = new Button
            {
                Text = "Refresh Views",
                Location = new Point(160, 5),
                Width = 100,
                Height = 25
            };
            refreshButton.Click += async (s, e) => await Task.Run(async () => await RefreshAllData());

            // Create Import Button
            importButton = new Button
            {
                Text = "Import from Excel",
                Location = new Point(280, 5), // Position after the refresh button
                Width = 120,
                Height = 25
            };
            importButton.Click += ImportButton_Click;

            exportButton = new Button
            {
                Text = "Export to Excel",
                Location = new Point(410, 5), // Position after the import button
                Width = 120,
                Height = 25
            };
            exportButton.Click += ExportButton_Click;

            startCountButton = new Button
            {
                Text = "Start Count",
                Location = new Point(10, 5),
                Width = 120,
                Height = 25
            };
            startCountButton.Click += StartCount_Click;

            timerLabel = new Label
            {
                Text = "Timer: Not Started",
                Location = new Point(740, 10),
                Width = 300, // Increased width for longer text
                AutoSize = true,
                Font = new Font(this.Font.FontFamily, 10, FontStyle.Bold)
            };

            // Add all controls to the panel
            topPanel.Controls.Add(refreshButton);
            topPanel.Controls.Add(importButton);
            topPanel.Controls.Add(exportButton);
            topPanel.Controls.Add(startCountButton);
            topPanel.Controls.Add(timerLabel);

            countdownTimer = new System.Windows.Forms.Timer
            {
                Interval = 1000 // 1 second
            };
            countdownTimer.Tick += CountdownTimer_Tick;

            // Setup DataGridView
            dataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.Fixed3D,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                AutoGenerateColumns = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                Margin = new Padding(0, 10, 0, 0)
            };

            // Add columns
            dataGridView.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn
                {
                    Name = "Title",
                    HeaderText = "Video Title",
                    Width = 300,
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "URL",
                    HeaderText = "Video URL",
                    Width = 300,
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "PublishedDate",
                    HeaderText = "Published Date",
                    Width = 120,
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "CurrentViews",
                    HeaderText = "Current Views",
                    Width = 150,
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Rank",
                    HeaderText = "Rank (1-10)",
                    Width = 100,
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Day0",
                    HeaderText = "Day 0",
                    Width = 100,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = DataGridViewContentAlignment.MiddleRight,
                        Format = "N0"
                    }
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Day1",
                    HeaderText = "Day 1",
                    Width = 100,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = DataGridViewContentAlignment.MiddleRight,
                        Format = "N0"
                    }
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Day2",
                    HeaderText = "Day 2",
                    Width = 100,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = DataGridViewContentAlignment.MiddleRight,
                        Format = "N0"
                    }
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Day3",
                    HeaderText = "Day 3",
                    Width = 100,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = DataGridViewContentAlignment.MiddleRight,
                        Format = "N0"
                    }
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Day4",
                    HeaderText = "Day 4",
                    Width = 100,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = DataGridViewContentAlignment.MiddleRight,
                        Format = "N0"
                    }
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Day5",
                    HeaderText = "Day 5",
                    Width = 100,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = DataGridViewContentAlignment.MiddleRight,
                        Format = "N0"
                    }
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Day6",
                    HeaderText = "Day 6",
                    Width = 100,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = DataGridViewContentAlignment.MiddleRight,
                        Format = "N0"
                    }
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Day7",
                    HeaderText = "Day 7",
                    Width = 100,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = DataGridViewContentAlignment.MiddleRight,
                        Format = "N0"
                    }
                }
            });

            // Setup layout
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            mainLayout.Controls.Add(topPanel, 0, 0);
            mainLayout.Controls.Add(dataGridView, 0, 1);

            this.Controls.Add(mainLayout);
        }

        private async Task LoadSampleData()
        {
            try
            {
                // Only load sample data if no videos have been imported
                if (videoUrls.Count == 0)
                {
                    videoUrls = new List<string>
            {
                "https://www.youtube.com/watch?v=wWOMLLJ-33c",
                "https://www.youtube.com/watch?v=1PAgV4yQm1M",
                "https://www.youtube.com/watch?v=cTwmL7Z30U8",
                "https://www.youtube.com/watch?v=tP9t3KYTy1Q",
                "https://www.youtube.com/watch?v=fw18VpHgytY",
                "https://www.youtube.com/watch?v=fg1G1_lSCpk",
                "https://www.youtube.com/watch?v=w9N-y1KE88o",
                "https://www.youtube.com/watch?v=aQnafaDQxvU",
                "https://www.youtube.com/watch?v=1gQ_mAzfLEU",
                "https://www.youtube.com/watch?v=tDAnoUIbMHY"
            };
                }

                await LoadVideoData();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading sample data: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task RefreshAllData()
        {
            try
            {
                refreshButton.Enabled = false;
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    string url = dataGridView.Rows[i].Cells[1].Value?.ToString();
                    if (!string.IsNullOrEmpty(url))
                    {
                        string videoId = ExtractVideoId(url);
                        await FetchAndUpdateVideoData(videoId, i);
                    }
                }
            }
            finally
            {
                refreshButton.Enabled = true;
            }
        }

        private async Task FetchAndUpdateVideoData(string videoId, int rowIndex)
        {
            try
            {
                string apiUrl = $"https://www.googleapis.com/youtube/v3/videos?id={videoId}&key={API_KEY}&part=statistics,snippet";
                var response = await httpClient.GetStringAsync(apiUrl);
                var json = JObject.Parse(response);

                var items = json["items"];
                if (items == null || !items.Any())
                {
                    MessageBox.Show($"No data found for video {videoId}");
                    return;
                }

                var videoData = items[0];
                var snippetData = videoData["snippet"];

                // Get video details
                string title = snippetData["title"].ToString();
                DateTime publishedAt = DateTime.Parse(snippetData["publishedAt"].ToString());
                long currentViews = long.Parse(videoData["statistics"]["viewCount"].ToString());

                // Calculate daily views for the last 7 days
                var dailyViews = new List<long>();
                for (int i = 0; i < 7; i++)
                {
                    DateTime date = DateTime.Today.AddDays(-i);
                    long views = SimulateHistoricalViews(currentViews, date);
                    dailyViews.Add(views);
                }

                // Calculate rank based on current views
                int rank = CalculateRank(currentViews);

                if (this.IsDisposed) return;

                this.Invoke((MethodInvoker)delegate
                {
                    if (rowIndex < dataGridView.Rows.Count)
                    {
                        DataGridViewRow row = dataGridView.Rows[rowIndex];

                        // Update basic information
                        row.Cells[0].Value = title;
                        row.Cells[1].Value = $"https://www.youtube.com/watch?v={videoId}";
                        row.Cells[2].Value = publishedAt.ToString("dd MMM yyyy");
                        row.Cells[3].Value = currentViews.ToString("N0");
                        row.Cells[4].Value = rank.ToString();

                        // Update daily views columns
                        for (int i = 0; i < 7; i++)
                        {
                            row.Cells[5 + i].Value = dailyViews[i];

                            // Color code daily views based on growth
                            if (i < 6)
                            {
                                long difference = dailyViews[i] - dailyViews[i + 1];
                                if (difference > 0)
                                {
                                    row.Cells[5 + i].Style.BackColor = Color.LightGreen;
                                    row.Cells[5 + i].Style.ForeColor = Color.DarkGreen;
                                }
                                else if (difference < 0)
                                {
                                    row.Cells[5 + i].Style.BackColor = Color.LightPink;
                                    row.Cells[5 + i].Style.ForeColor = Color.DarkRed;
                                }
                                else
                                {
                                    row.Cells[5 + i].Style.BackColor = SystemColors.Window;
                                    row.Cells[5 + i].Style.ForeColor = SystemColors.ControlText;
                                }
                            }
                        }

                        // Style for recent videos
                        TimeSpan videoAge = DateTime.Now - publishedAt;
                        if (videoAge.TotalDays <= 30)
                        {
                            row.Cells[2].Style.BackColor = Color.LightYellow;
                            row.Cells[2].Style.Font = new Font(dataGridView.Font, FontStyle.Bold);
                        }
                        else
                        {
                            row.Cells[2].Style.BackColor = SystemColors.Window;
                            row.Cells[2].Style.Font = dataGridView.Font;
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                if (!this.IsDisposed)
                {
                    this.Invoke((MethodInvoker)delegate
                    {
                        MessageBox.Show($"Error fetching data for video {videoId}: {ex.Message}", "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }
            }
        }

        private string ExtractVideoId(string url)
        {
            try
            {
                if (url.Contains("youtube.com/watch?v="))
                {
                    return url.Split('=')[1].Split('&')[0];
                }
                else if (url.Contains("youtu.be/"))
                {
                    return url.Split('/').Last();
                }
            }
            catch { }
            return string.Empty;
        }

        private long SimulateHistoricalViews(long currentViews, DateTime targetDate)
        {
            if (targetDate.Date == DateTime.Today.Date)
                return currentViews;

            // If target date is in the future, return current views
            if (targetDate.Date > DateTime.Today.Date)
                return currentViews;

            // Calculate days difference
            int daysDifference = (DateTime.Today.Date - targetDate.Date).Days;

            // Assuming an average daily growth rate between 0.1% to 0.5%
            Random rand = new Random(targetDate.GetHashCode()); // Use consistent seed for same date
            double dailyGrowthRate = 0.001 + (rand.NextDouble() * 0.004); // 0.1% to 0.5%

            // Calculate historical views using compound reduction
            double totalReductionFactor = Math.Pow(1 + dailyGrowthRate, daysDifference);
            return (long)(currentViews / totalReductionFactor);
        }

        private int CalculateRank(long views)
        {
            if (views < 1000) return 1;
            else if (views < 10000) return 2;
            else if (views < 50000) return 3;
            else if (views < 100000) return 4;
            else if (views < 500000) return 5;
            else if (views < 1000000) return 6;
            else if (views < 5000000) return 7;
            else if (views < 10000000) return 8;
            else if (views < 50000000) return 9;
            else return 10;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            httpClient?.Dispose();
        }

        // Add these new methods for Excel import functionality
        private void ImportButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                openFileDialog.Title = "Select Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ImportExcelFile(openFileDialog.FileName);
                }
            }
        }

        private void ImportExcelFile(string filePath)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });

                        // Get the first worksheet
                        DataTable dataTable = result.Tables[0];
                        ProcessExcelData(dataTable);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error importing Excel file: {ex.Message}", "Import Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ProcessExcelData(DataTable dataTable)
        {
            try
            {
                // Clear existing data
                videoUrls.Clear();
                dataGridView.Rows.Clear();

                // Look for a column containing YouTube URLs
                var urlColumn = FindYouTubeUrlColumn(dataTable);
                if (urlColumn == null)
                {
                    MessageBox.Show("No column with YouTube URLs found in the Excel file.\n" +
                                  "Please ensure your Excel file has a column with YouTube URLs.",
                                  "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Process each row
                foreach (DataRow row in dataTable.Rows)
                {
                    string url = row[urlColumn].ToString().Trim();
                    if (!string.IsNullOrEmpty(url) && IsValidYouTubeUrl(url))
                    {
                        videoUrls.Add(url);
                    }
                }

                if (videoUrls.Count == 0)
                {
                    MessageBox.Show("No valid YouTube URLs found in the Excel file.",
                                  "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                LoadVideoData();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing Excel data: {ex.Message}", "Processing Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string FindYouTubeUrlColumn(DataTable dataTable)
        {
            // Look for common column names that might contain YouTube URLs
            string[] possibleColumnNames = new[]
            {
        "url", "link", "youtube", "video", "youtube url", "video url",
        "youtube link", "video link"
    };

            // First try exact match (case-insensitive)
            foreach (DataColumn column in dataTable.Columns)
            {
                if (possibleColumnNames.Contains(column.ColumnName.ToLower()))
                {
                    return column.ColumnName;
                }
            }

            // Then try contains match
            foreach (DataColumn column in dataTable.Columns)
            {
                foreach (string name in possibleColumnNames)
                {
                    if (column.ColumnName.ToLower().Contains(name))
                    {
                        return column.ColumnName;
                    }
                }
            }

            // If no matching column name found, look for a column containing YouTube URLs
            foreach (DataColumn column in dataTable.Columns)
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    string value = row[column].ToString();
                    if (!string.IsNullOrEmpty(value) && IsValidYouTubeUrl(value))
                    {
                        return column.ColumnName;
                    }
                }
            }

            return null;
        }

        private bool IsValidYouTubeUrl(string url)
        {
            if (string.IsNullOrEmpty(url)) return false;

            return url.Contains("youtube.com/watch?v=") || url.Contains("youtu.be/");
        }

        private async Task LoadVideoData()
        {
            try
            {
                dataGridView.Rows.Clear();
                foreach (string url in videoUrls)
                {
                    string videoId = ExtractVideoId(url);
                    if (!string.IsNullOrEmpty(videoId))
                    {
                        int rowIndex = dataGridView.Rows.Add(
                            "Loading...", // Title
                            url,         // URL
                            "Loading...", // Published Date
                            "Loading...", // Current Views
                            "Loading...", // View Range
                            "-",         // Rank
                            "Loading...", // Day 1
                            "Loading...", // Day 2
                            "Loading...", // Day 3
                            "Loading...", // Day 4
                            "Loading...", // Day 5
                            "Loading...", // Day 6
                            "Loading..."  // Day 7
                        );
                        await FetchAndUpdateVideoData(videoId, rowIndex);
                    }
                }

                MessageBox.Show($"Successfully imported {videoUrls.Count} videos.", "Import Complete",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading video data: {ex.Message}", "Loading Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView.Rows.Count == 0)
                {
                    MessageBox.Show("No data to export.", "Export Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.Title = "Save Excel File";
                    saveFileDialog.DefaultExt = "xlsx";
                    saveFileDialog.FileName = $"YouTube_Stats_{DateTime.Now:yyyy-MM-dd}";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        ExportToExcel(saveFileDialog.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during export: {ex.Message}", "Export Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportToExcel(string filePath)
        {
            try
            {
                // Create a new Excel package
                using (var package = new OfficeOpenXml.ExcelPackage())
                {
                    // Add a new worksheet to the package
                    var worksheet = package.Workbook.Worksheets.Add("YouTube Stats");

                    // Add headers with formatting
                    for (int i = 0; i < dataGridView.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dataGridView.Columns[i].HeaderText;
                        worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                        worksheet.Cells[1, i + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                        worksheet.Cells[1, i + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    }

                    // Add data
                    for (int i = 0; i < dataGridView.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView.Columns.Count; j++)
                        {
                            var cell = worksheet.Cells[i + 2, j + 1];
                            var value = dataGridView.Rows[i].Cells[j].Value;

                            // Format based on column type
                            if (j >= 5) // Day columns (numeric)
                            {
                                if (long.TryParse(value?.ToString().Replace(",", ""), out long numericValue))
                                {
                                    cell.Value = numericValue;
                                    cell.Style.Numberformat.Format = "#,##0";
                                }
                                else
                                {
                                    cell.Value = value;
                                }
                            }
                            else if (j == 3) // Current Views column
                            {
                                if (long.TryParse(value?.ToString().Replace(",", ""), out long numericValue))
                                {
                                    cell.Value = numericValue;
                                    cell.Style.Numberformat.Format = "#,##0";
                                }
                                else
                                {
                                    cell.Value = value;
                                }
                            }
                            else
                            {
                                cell.Value = value;
                            }

                            // Add border
                            cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        }
                    }

                    // Auto-fit columns
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    // Set minimum width for numeric columns
                    for (int i = 6; i <= dataGridView.Columns.Count; i++)
                    {
                        if (worksheet.Column(i).Width < 12)
                            worksheet.Column(i).Width = 12;
                    }

                    // Save the file
                    var file = new FileInfo(filePath);
                    package.SaveAs(file);

                    MessageBox.Show("Data exported successfully!", "Export Complete",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting to Excel: {ex.Message}", "Export Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void StartCount_Click(object sender, EventArgs e)
        {
            try
            {
                var result = MessageBox.Show(
                    "This will start a 7-day counting process. Continue?",
                    "Start Count",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    // Disable buttons during count
                    startCountButton.Enabled = false;
                    importButton.Enabled = false;

                    // Reset counters
                    currentDay = 0;
                    minutesRemaining = 168 * 60;
                    lastRefreshTime = DateTime.Now;

                    // Initial refresh and update Day 0
                    await RefreshAndUpdateDay();

                    // Start timer
                    countdownTimer.Start();
                    UpdateTimerDisplay();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting count: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CountdownTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                TimeSpan elapsed = DateTime.Now - lastRefreshTime;

                // Check if 24 hours have passed
                if (elapsed.TotalHours >= 24)
                {
                    currentDay++;
                    if (currentDay <= 7)
                    {
                        // Trigger refresh and update for the new day
                        RefreshAndUpdateDay().ConfigureAwait(false);
                        lastRefreshTime = DateTime.Now;
                    }
                }

                // Update remaining time in minutes
                minutesRemaining = (168 * 60) - (int)Math.Floor(
                    (DateTime.Now - lastRefreshTime).TotalMinutes + (currentDay * 24 * 60));

                // Update timer display
                UpdateTimerDisplay();

                // Check if complete
                if (minutesRemaining <= 0 || currentDay > 7)
                {
                    CompleteCount();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error in timer: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                CompleteCount();
            }
        }

        private async Task RefreshAndUpdateDay()
        {
            try
            {
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    string url = dataGridView.Rows[i].Cells[1].Value?.ToString();
                    if (!string.IsNullOrEmpty(url))
                    {
                        string videoId = ExtractVideoId(url);
                        await FetchAndUpdateDayViews(videoId, i, currentDay);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating day {currentDay}: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task FetchAndUpdateDayViews(string videoId, int rowIndex, int day)
        {
            try
            {
                string apiUrl = $"https://www.googleapis.com/youtube/v3/videos?id={videoId}&key={API_KEY}&part=statistics";
                var response = await httpClient.GetStringAsync(apiUrl);
                var json = JObject.Parse(response);

                var items = json["items"];
                if (items == null || !items.Any())
                {
                    MessageBox.Show($"No data found for video {videoId}");
                    return;
                }

                var videoData = items[0];
                long currentViews = long.Parse(videoData["statistics"]["viewCount"].ToString());

                if (this.IsDisposed) return;

                this.Invoke((MethodInvoker)delegate
                {
                    if (rowIndex < dataGridView.Rows.Count)
                    {
                        var row = dataGridView.Rows[rowIndex];
                        row.Cells[5 + day].Value = currentViews.ToString("N0");

                        // Color code the cell based on change from previous day
                        if (day > 0)
                        {
                            var previousViews = long.Parse(row.Cells[5 + day - 1].Value.ToString().Replace(",", ""));
                            var difference = currentViews - previousViews;

                            if (difference > 0)
                            {
                                row.Cells[5 + day].Style.BackColor = Color.LightGreen;
                                row.Cells[5 + day].Style.ForeColor = Color.DarkGreen;
                            }
                            else if (difference < 0)
                            {
                                row.Cells[5 + day].Style.BackColor = Color.LightPink;
                                row.Cells[5 + day].Style.ForeColor = Color.DarkRed;
                            }
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Error updating views for day {day}: {ex.Message}");
            }
        }

        private void UpdateTimerDisplay()
        {
            int days = minutesRemaining / (24 * 60);
            int hours = (minutesRemaining % (24 * 60)) / 60;
            int minutes = minutesRemaining % 60;

            string timeDisplay = string.Format(
                "Timer: {0}d {1}h {2}m remaining (Day {3}/7)",
                days,
                hours.ToString("00"),
                minutes.ToString("00"),
                currentDay
            );

            // Add progress percentage
            double totalMinutes = 168.0 * 60;
            double remainingMinutes = minutesRemaining;
            int progressPercent = (int)Math.Round(((totalMinutes - remainingMinutes) / totalMinutes) * 100);

            timerLabel.Text = $"{timeDisplay} - {progressPercent}% Complete";

            // Change color based on remaining time
            if (days <= 1)  // Less than 1 day remaining
            {
                timerLabel.ForeColor = Color.Red;
            }
            else if (days <= 2)  // Less than 2 days remaining
            {
                timerLabel.ForeColor = Color.OrangeRed;
            }
            else
            {
                timerLabel.ForeColor = Color.Black;
            }
        }

        private void CompleteCount()
        {
            countdownTimer.Stop();
            startCountButton.Enabled = true;
            importButton.Enabled = true;
            timerLabel.Text = "Count Complete - 100% Done";
            timerLabel.ForeColor = Color.Green;
            MessageBox.Show("View counting process complete!", "Complete",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

    }
}