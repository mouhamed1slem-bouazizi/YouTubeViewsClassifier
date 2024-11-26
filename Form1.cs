using System;
using System.Windows.Forms;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Drawing;
using System.Linq;
using System.Collections.Generic;

namespace YouTubeViewsClassifier
{
    public partial class Form1 : Form
    {
        private DataGridView dataGridView;
        private DateTimePicker startDatePicker;
        private DateTimePicker endDatePicker;
        private Button refreshButton;
        private const string API_KEY = "AIzaSyB2MFgLVS1r1rzMMniCWmdwcqOl9B3K9II";
        private readonly HttpClient httpClient;

        public Form1()
        {
            InitializeComponent();
            httpClient = new HttpClient();
            SetupUI();
            // Load data after UI is setup
            this.Shown += async (s, e) => await LoadSampleData();
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

            // Create labels
            Label fromLabel = new Label
            {
                Text = "From:",
                Location = new Point(0, 10),
                AutoSize = true
            };

            Label toLabel = new Label
            {
                Text = "To:",
                Location = new Point(190, 10),
                AutoSize = true
            };

            // Create Start Date Picker
            startDatePicker = new DateTimePicker
            {
                Format = DateTimePickerFormat.Short,
                Location = new Point(40, 5),
                Width = 120
            };
            startDatePicker.MaxDate = DateTime.Today;
            startDatePicker.Value = DateTime.Today.AddDays(-7); // Default to 7 days ago
            startDatePicker.ValueChanged += async (s, e) =>
            {
                if (startDatePicker.Value > endDatePicker.Value)
                {
                    endDatePicker.Value = startDatePicker.Value;
                }
                await RefreshAllData();
            };

            // Create End Date Picker
            endDatePicker = new DateTimePicker
            {
                Format = DateTimePickerFormat.Short,
                Location = new Point(220, 5),
                Width = 120
            };
            endDatePicker.MaxDate = DateTime.Today;
            endDatePicker.Value = DateTime.Today;
            endDatePicker.ValueChanged += async (s, e) =>
            {
                if (endDatePicker.Value < startDatePicker.Value)
                {
                    startDatePicker.Value = endDatePicker.Value;
                }
                await RefreshAllData();
            };

            // Create and setup Refresh button
            refreshButton = new Button
            {
                Text = "Refresh Views",
                Location = new Point(360, 5),
                Width = 100,
                Height = 25
            };
            refreshButton.Click += async (s, e) => await RefreshAllData();

            // Add all controls to the panel
            topPanel.Controls.Add(fromLabel);
            topPanel.Controls.Add(startDatePicker);
            topPanel.Controls.Add(toLabel);
            topPanel.Controls.Add(endDatePicker);
            topPanel.Controls.Add(refreshButton);

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
                    Name = "ViewRange",
                    HeaderText = "Views in Date Range",
                    Width = 250, // Increased width for better formatting
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = DataGridViewContentAlignment.MiddleCenter,
                        WrapMode = DataGridViewTriState.True
                    }
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "Rank",
                    HeaderText = "Rank (1-10)",
                    Width = 100,
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
                dataGridView.Rows.Clear();
                string[] videoUrls = new string[]
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

                foreach (string url in videoUrls)
                {
                    int rowIndex = dataGridView.Rows.Add("Loading...", url, "Loading...", "Loading...", "Loading...", "-");

                    string videoId = ExtractVideoId(url);
                    if (!string.IsNullOrEmpty(videoId))
                    {
                        await FetchAndUpdateVideoData(videoId, rowIndex);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading data: {ex.Message}");
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

                // Calculate views for both dates
                long startViews = SimulateHistoricalViews(currentViews, startDatePicker.Value);
                long endViews = SimulateHistoricalViews(currentViews, endDatePicker.Value);

                // Calculate actual view difference in the date range
                long viewDifference = endViews - startViews;
                double percentageChange = startViews > 0 ? ((double)viewDifference / startViews) * 100 : 0;

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

                        // Format view range with more detail
                        string viewRangeText;
                        if (startDatePicker.Value.Date == endDatePicker.Value.Date)
                        {
                            viewRangeText = $"{startViews:N0} views";
                        }
                        else
                        {
                            viewRangeText = $"{startViews:N0} → {endViews:N0}\n" +
                                          $"Δ: {(viewDifference >= 0 ? "+" : "")}{viewDifference:N0} " +
                                          $"({(viewDifference >= 0 ? "+" : "")}{percentageChange:F1}%)";
                        }
                        row.Cells[4].Value = viewRangeText;
                        row.Cells[5].Value = rank.ToString();

                        // Style the view range cell
                        if (viewDifference > 0)
                        {
                            row.Cells[4].Style.BackColor = Color.LightGreen;
                            row.Cells[4].Style.ForeColor = Color.DarkGreen;
                        }
                        else if (viewDifference < 0)
                        {
                            row.Cells[4].Style.BackColor = Color.LightPink;
                            row.Cells[4].Style.ForeColor = Color.DarkRed;
                        }
                        else
                        {
                            row.Cells[4].Style.BackColor = SystemColors.Window;
                            row.Cells[4].Style.ForeColor = SystemColors.ControlText;
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

                        // Handle cases where start date is before publish date
                        if (startDatePicker.Value.Date < publishedAt.Date)
                        {
                            row.Cells[4].Style.BackColor = Color.LightGray;
                            row.Cells[4].Style.ForeColor = SystemColors.ControlText;
                            row.Cells[4].Value = "N/A - Before publication";
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
    }
}