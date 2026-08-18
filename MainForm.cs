using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Drawing;


namespace TusurOcenkiGUI
{
    public partial class MainForm : Form
    {
        private CancellationTokenSource cts;
        private TextBox txtFilePath;
        private NumericUpDown numSurname, numName, numGroup, numPrev, numLast, numMean, numDelay;
        private Button btnBrowse, btnStart, btnStop;
        private ProgressBar progressBar;
        private RichTextBox logBox;
        private Label lblProgress;
        private DateTime startTime;
        private int processedCount = 0;
        private const string SettingsFile = "settings.txt";
        private Button btnClearLog;
        private bool isPaused = false;
        private Label lbl5, lbl45, lbl4, lbl35, lbl3, lbl25, lblLow;
        private CheckBox chkAutoScroll;
        private TabControl tabControl;
        private Dictionary<string, ListView> gradeLists = new Dictionary<string, ListView>();
        private NumericUpDown numRoom;
        private bool isProcessing = false;

        public MainForm()
        {
            InitializeComponentCustom();
        }

        private void InitializeComponentCustom()
        {
            this.Text = "ТУСУР Парсер оценок - v4.0 / ©PDG_572-1";
            this.Size = new System.Drawing.Size(920, 720);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);

            Label lblFile = new Label { Text = "Excel файл:", Location = new System.Drawing.Point(20, 20), AutoSize = true };
            txtFilePath = new TextBox { Location = new System.Drawing.Point(120, 18), Size = new System.Drawing.Size(580, 23), ReadOnly = true };
            btnBrowse = new Button { Text = "Выбрать файл", Location = new System.Drawing.Point(710, 17), Size = new System.Drawing.Size(110, 30) };
            btnBrowse.Click += BtnBrowse_Click;

            Label lblCols = new Label { Text = "Номера столбцов(A-1, B-2 и тд)", Location = new System.Drawing.Point(20, 70), AutoSize = true };

            numSurname = CreateNumeric(20, 115, "Фамилия");
            numName = CreateNumeric(235, 115, "Имя");
            numGroup = CreateNumeric(450, 115, "Группа");
            numPrev = CreateNumeric(20, 165, "Предпоследняя ЭС");
            numLast = CreateNumeric(235, 165, "Последняя ЭС");
            numMean = CreateNumeric(450, 165, "Средний балл");
            numRoom = CreateNumeric(680, 115, "Комната (необязательно)");

            Label lblDelay = new Label { Text = "Задержка (мс):", Location = new System.Drawing.Point(20, 200), AutoSize = true, Font = new System.Drawing.Font("Segoe UI", 10F) };
            numDelay = new NumericUpDown { Location = new System.Drawing.Point(160, 200), Size = new System.Drawing.Size(100, 28), Minimum = 0, Maximum = 5000, Value = 100, Increment = 100, Font = new System.Drawing.Font("Segoe UI", 10F) };

            btnStart = new Button { Text = "Запустить обработку", Location = new System.Drawing.Point(20, 245), Size = new System.Drawing.Size(250, 50), BackColor = System.Drawing.Color.LightGreen, Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Bold) };
            btnStart.Click += BtnStart_Click;

            btnStop = new Button { Text = "■ Остановить", Location = new System.Drawing.Point(280, 245), Size = new System.Drawing.Size(150, 50), BackColor = System.Drawing.Color.Red, Enabled = false, Font = new System.Drawing.Font("Segoe UI", 10F) };
            btnStop.Click += BtnStop_Click;
            this.Controls.Add(btnStop);

            btnClearLog = new Button
            {
                Text = "Очистить лог",
                Location = new System.Drawing.Point(20, 620),
                Size = new System.Drawing.Size(150, 40),
                Font = new System.Drawing.Font("Segoe UI", 10F)
            };
            btnClearLog.Click += BtnClearLog_Click;

            chkAutoScroll = new CheckBox
            {
                Text = "Авто-скролл",
                Location = new System.Drawing.Point(200, 615),
                Size = new System.Drawing.Size(180, 60),
                Checked = true,
                Font = new System.Drawing.Font("Segoe UI", 11F)
            };
            this.Controls.Add(chkAutoScroll);

            lbl5 = new Label { Text = "5.0: 0", Location = new System.Drawing.Point(400, 615), AutoSize = true, Font = new System.Drawing.Font("Segoe UI", 12F) };
            lbl45 = new Label { Text = "4.5-4.99: 0", Location = new System.Drawing.Point(400, 640), AutoSize = true, Font = new System.Drawing.Font("Segoe UI", 12F) };
            lbl4 = new Label { Text = "4.0-4.49: 0", Location = new System.Drawing.Point(530, 615), AutoSize = true, Font = new System.Drawing.Font("Segoe UI", 12F) };
            lbl35 = new Label { Text = "3.5-3.99: 0", Location = new System.Drawing.Point(530, 640), AutoSize = true, Font = new System.Drawing.Font("Segoe UI", 12F) };
            lbl3 = new Label { Text = "3.0-3.49: 0", Location = new System.Drawing.Point(660, 615), AutoSize = true, Font = new System.Drawing.Font("Segoe UI", 12F) };
            lbl25 = new Label { Text = "2.5-2.99: 0", Location = new System.Drawing.Point(660, 640), AutoSize = true, Font = new System.Drawing.Font("Segoe UI", 12F) };
            lblLow = new Label { Text = "< 2.5: 0", Location = new System.Drawing.Point(790, 640), AutoSize = true, Font = new System.Drawing.Font("Segoe UI", 12F) };

            this.Controls.Add(lbl5);
            this.Controls.Add(lbl45);
            this.Controls.Add(lbl4);
            this.Controls.Add(lbl35);
            this.Controls.Add(lbl3);
            this.Controls.Add(lbl25);
            this.Controls.Add(lblLow);

            progressBar = new ProgressBar { Location = new System.Drawing.Point(20, 300), Size = new System.Drawing.Size(790, 25) };

            lblProgress = new Label
            {
                Location = new System.Drawing.Point(260, 270),
                Size = new System.Drawing.Size(790, 20),
                Text = "Готов к запуску",
                Font = new System.Drawing.Font("Segoe UI", 10F),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };

            tabControl = new TabControl
            {
                Location = new System.Drawing.Point(20, 335),
                Size = new System.Drawing.Size(800, 270)
            };

            TabPage tabMain = new TabPage("Главная");

            Panel logPanel = new Panel
            {
                Location = new System.Drawing.Point(0, 0),
                Size = new System.Drawing.Size(800, 270),
                BackColor = System.Drawing.Color.FromArgb(245, 245, 245),
                BorderStyle = BorderStyle.FixedSingle
            };

            logBox = new RichTextBox
            {
                Location = new System.Drawing.Point(5, 5),
                Size = new System.Drawing.Size(790, 260),
                Multiline = true,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 11f),
                BackColor = System.Drawing.Color.FromArgb(245, 245, 245),
                ForeColor = System.Drawing.Color.Black,
                BorderStyle = BorderStyle.None
            };

            logPanel.Controls.Add(logBox);
            tabMain.Controls.Add(logPanel);
            tabControl.TabPages.Add(tabMain);

            CreateGradeTab("5.0", Color.Green);
            CreateGradeTab("4.5-4.99", Color.LimeGreen);
            CreateGradeTab("4.0-4.49", Color.YellowGreen);
            CreateGradeTab("3.5-3.99", Color.Orange);
            CreateGradeTab("3.0-3.49", Color.OrangeRed);
            CreateGradeTab("2.5-2.99", Color.Red);
            CreateGradeTab("< 2.5", Color.DarkRed);

            this.Controls.Add(tabControl);

            this.Controls.Add(lblFile);
            this.Controls.Add(txtFilePath);
            this.Controls.Add(btnBrowse);
            this.Controls.Add(lblCols);
            this.Controls.Add(numSurname);
            this.Controls.Add(numName);
            this.Controls.Add(numGroup);
            this.Controls.Add(numPrev);
            this.Controls.Add(numLast);
            this.Controls.Add(numMean);
            this.Controls.Add(lblDelay);
            this.Controls.Add(numDelay);
            this.Controls.Add(btnStart);
            this.Controls.Add(progressBar);
            this.Controls.Add(lblProgress);
            this.Controls.Add(btnClearLog);

            txtFilePath.AllowDrop = true;
            txtFilePath.DragEnter += TxtFilePath_DragEnter;
            txtFilePath.DragDrop += TxtFilePath_DragDrop;
            txtFilePath.DragLeave += TxtFilePath_DragLeave;

            this.AllowDrop = true;
            this.DragEnter += TxtFilePath_DragEnter;
            this.DragDrop += TxtFilePath_DragDrop;
            this.DragLeave += TxtFilePath_DragLeave;
        }

        private NumericUpDown CreateNumeric(int x, int y, string labelText)
        {
            Label lbl = new Label { Text = labelText + ":", Location = new System.Drawing.Point(x, y - 18), AutoSize = true };
            NumericUpDown num = new NumericUpDown { Location = new System.Drawing.Point(x, y), Size = new System.Drawing.Size(200, 40), Value = 0, Minimum = 0, Maximum = 50 };
            this.Controls.Add(lbl);
            this.Controls.Add(num);
            return num;
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel файлы|*.xlsx;*.xls";
                if (ofd.ShowDialog() == DialogResult.OK)
                    txtFilePath.Text = ofd.FileName;
            }
        }

        private async void BtnStart_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFilePath.Text))
            {
                MessageBox.Show("Выберите Excel файл!");
                return;
            }

            if (!ValidateColumns())
                return;

            LockControls(true);

            if (btnStart != null) btnStart.Enabled = false;
            if (btnStop != null)
            {
                btnStop.Enabled = true;
                btnStop.Text = "■ Остановить";
            }

            logBox.Clear();
            cts = new CancellationTokenSource();

            try
            {
                await ProcessFile(cts.Token);
                Log("=== Завершено успешно ===");
            }
            catch (OperationCanceledException)
            {
                Log("Обработка остановлена пользователем.");
            }
            catch (Exception ex)
            {
                Log("Ошибка: " + ex.Message);
            }
            finally
            {
                ResetUIAfterProcessing();
            }
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            if (cts == null) return;

            if (!isPaused)
            {
                isPaused = true;
                if (btnStop != null) btnStop.Text = "▶ Возобновить";
                Log("Обработка приостановлена...");
            }
            else
            {
                isPaused = false;
                if (btnStop != null) btnStop.Text = "■ Остановить";
                Log("Обработка возобновлена...");
            }
        }

        private void ResetUIAfterProcessing()
        {
            if (progressBar != null) progressBar.Value = 0;
            if (lblProgress != null) lblProgress.Text = "Готов к запуску";
            if (btnStart != null) btnStart.Enabled = true;
            if (btnStop != null) btnStop.Enabled = false;
            isPaused = false;
            if (btnStop != null) btnStop.Text = "■ Остановить";

            LockControls(false);
            this.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
        }

        private void Log(string message, Color? color = null)
        {
            if (logBox.InvokeRequired)
            {
                logBox.Invoke(new Action(() => AppendColoredText(message, color)));
            }
            else
            {
                AppendColoredText(message, color);
            }
        }

        private void AppendColoredText(string message, Color? color)
        {
            logBox.SelectionStart = logBox.TextLength;
            logBox.SelectionLength = 0;

            if (color.HasValue)
                logBox.SelectionColor = color.Value;
            else
                logBox.SelectionColor = Color.Black;

            logBox.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + message + "\r\n");

            logBox.SelectionColor = logBox.ForeColor;

            if (chkAutoScroll != null && chkAutoScroll.Checked)
            {
                logBox.SelectionStart = logBox.TextLength;
                logBox.ScrollToCaret();
            }
        }

        private async Task ProcessFile(CancellationToken token)
        {
            startTime = DateTime.Now;
            processedCount = 0;
            int totalStudents = 0;
            int count5 = 0;
            int count45_499 = 0;
            int count4_449 = 0;
            int count35_399 = 0;
            int count3_349 = 0;
            int count25_299 = 0;
            int countLow = 0;

            using var client = new HttpClient();
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");

            using (var workbook = new XLWorkbook(txtFilePath.Text))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RangeUsed().RowsUsed().Skip(1).ToList();
                int total = rows.Count;

                foreach (var excelRow in rows)
                {
                    token.ThrowIfCancellationRequested();

                    while (isPaused)
                    {
                        await Task.Delay(200, token);
                    }

                    string surname = excelRow.Cell((int)numSurname.Value).GetFormattedString().Trim();
                    string name = excelRow.Cell((int)numName.Value).GetFormattedString().Trim();
                    string group = excelRow.Cell((int)numGroup.Value).GetFormattedString().Trim();

                    if (group.Length == 2) group = "0" + group;

                    Log($"{surname} {name} {group} → ");

                    double prevMark = 0;
                    double lastMark = 0;
                    double meanMark = 0;

                    try
                    {
                        using var search = await client.GetAsync($"https://ocenka.tusur.ru/student_search?utf8=%E2%9C%93&surname={surname}&name={name}&group={group}&commit=%D0%9D%D0%B0%D0%B9%D1%82%D0%B8", token);
                        var content = await search.Content.ReadAsStringAsync();

                        if (content.Length > 150)
                        {
                            Log("Не найден", Color.Red);
                            processedCount++;
                            UpdateProgress(processedCount, total);
                            continue;
                        }

                        string link = content.Split("\"")[1];

                        using var result = await client.GetAsync(link, token);
                        content = await result.Content.ReadAsStringAsync();

                        int index = content.IndexOf("&quot;context_id&quot;:");
                        string context = content.Substring(index + "&quot;context_id&quot;:".Length).Split(',')[0];

                        using var response = await client.GetAsync($"https://ocenka.tusur.ru/api/students/{context}/statistics?context_id={context}&context_type=student&kinds%5B%5D=exam_session&kinds%5B%5D=kt&role=student_search", token);
                        string json = await response.Content.ReadAsStringAsync();

                        JObject obj = JObject.Parse(json);
                        JArray values = (JArray)obj["values"];
                        JArray labels = (JArray)obj["labels"];

                        int sessionCount = 0;
                        double[] marks = new double[2];

                        for (int j = labels.Count - 1; j >= 0 && sessionCount < 2; j--)
                        {
                            if (labels[j].ToString().Contains("ЭС"))
                            {
                                marks[sessionCount] = (double)values[j];
                                sessionCount++;
                            }
                        }

                        if (sessionCount == 2)
                        {
                            lastMark = marks[0];
                            prevMark = marks[1];
                            meanMark = (lastMark + prevMark) / 2.0;
                        }

                        string groupStr = excelRow.Cell((int)numGroup.Value).GetFormattedString().Trim();
                        string roomStr = "-";

                        if (numRoom.Value >= 1)
                        {
                            roomStr = excelRow.Cell((int)numRoom.Value).GetFormattedString().Trim();
                        }

                        string gradeKey = GetGradeKey(meanMark);
                        if (gradeLists.ContainsKey(gradeKey))
                        {
                            ListViewItem item = new ListViewItem(new string[]
                            {
                                surname,
                                name,
                                groupStr,
                                roomStr,
                                meanMark.ToString("F2")
                            });
                            gradeLists[gradeKey].Items.Add(item);
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch
                    {
                        Log("Ошибка", Color.Red);
                    }

                    excelRow.Cell((int)numPrev.Value).Value = prevMark;
                    excelRow.Cell((int)numLast.Value).Value = lastMark;
                    excelRow.Cell((int)numMean.Value).Value = meanMark;

                    // Обновлённая статистика
                    if (meanMark >= 5.0) count5++;
                    else if (meanMark >= 4.5) count45_499++;
                    else if (meanMark >= 4.0) count4_449++;
                    else if (meanMark >= 3.5) count35_399++;
                    else if (meanMark >= 3.0) count3_349++;
                    else if (meanMark >= 2.5) count25_299++;
                    else countLow++;

                    totalStudents++;

                    Log($"Готово (Средний: {meanMark:F2})", Color.Green);

                    UpdateStatLabels(count5, count45_499, count4_449, count35_399, count3_349, count25_299, countLow);

                    workbook.Save();

                    processedCount++;
                    UpdateProgress(processedCount, total);

                    if ((int)numDelay.Value > 0)
                        await Task.Delay((int)numDelay.Value, token);
                }
            }

            Log("");
            Log("================== СТАТИСТИКА ==================");
            Log($"Всего обработано студентов: {totalStudents}");
            Log("");
            Log($"5.0                  : {count5}");
            Log($"4.5 - 4.99           : {count45_499}");
            Log($"4.0 - 4.49           : {count4_449}");
            Log($"3.5 - 3.99           : {count35_399}");
            Log($"3.0 - 3.49           : {count3_349}");
            Log($"2.5 - 2.99           : {count25_299}");
            Log($"< 2.5                : {countLow}");
            Log("==================================================");

            UpdateStatLabels(count5, count45_499, count4_449, count35_399, count3_349, count25_299, countLow);
        }

        private void UpdateProgress(int processed, int total)
        {
            if (total == 0) return;

            int percent = (int)((double)processed / total * 100);
            TimeSpan elapsed = DateTime.Now - startTime;
            double avgTimePerStudent = processed > 0 ? elapsed.TotalSeconds / processed : 0;
            double remainingSeconds = (total - processed) * avgTimePerStudent;

            string remainingText = remainingSeconds > 60
                ? $"{(int)remainingSeconds / 60} мин {(int)remainingSeconds % 60} сек"
                : $"{(int)remainingSeconds} сек";

            if (progressBar.InvokeRequired)
            {
                progressBar.Invoke(new Action(() =>
                {
                    progressBar.Value = percent;
                    if (lblProgress != null)
                        lblProgress.Text = $"Обработано {processed} из {total} ({percent}%) | Осталось ~{remainingText}";
                }));
            }
            else
            {
                progressBar.Value = percent;
                if (lblProgress != null)
                    lblProgress.Text = $"Обработано {processed} из {total} ({percent}%) | Осталось ~{remainingText}";
            }
        }

        private void UpdateStatLabels(int c5, int c45, int c4, int c35, int c3, int c25, int cLow)
        {
            if (lbl5 != null) lbl5.Text = $"5.0: {c5}";
            if (lbl45 != null) lbl45.Text = $"4.5-4.99: {c45}";
            if (lbl4 != null) lbl4.Text = $"4.0-4.49: {c4}";
            if (lbl35 != null) lbl35.Text = $"3.5-3.99: {c35}";
            if (lbl3 != null) lbl3.Text = $"3.0-3.49: {c3}";
            if (lbl25 != null) lbl25.Text = $"2.5-2.99: {c25}";
            if (lblLow != null) lblLow.Text = $"< 2.5: {cLow}";
        }

        private void TxtFilePath_DragEnter(object sender, DragEventArgs e)
        {
            if (isProcessing)
            {
                e.Effect = DragDropEffects.None;
                this.BackColor = System.Drawing.Color.FromArgb(255, 180, 180); // Красный
                return;
            }

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
                this.BackColor = System.Drawing.Color.FromArgb(220, 240, 255); // Голубой
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void TxtFilePath_DragLeave(object sender, EventArgs e)
        {
            this.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
        }

        private void TxtFilePath_DragDrop(object sender, DragEventArgs e)
        {
            this.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);

            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Length > 0)
            {
                string filePath = files[0];
                if (filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                    filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    txtFilePath.Text = filePath;
                    Log("Файл загружен через Drag & Drop: " + Path.GetFileName(filePath));
                }
                else
                {
                    MessageBox.Show("Поддерживаются только файлы .xlsx и .xls", "Неверный формат", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private bool ValidateColumns()
        {
            var columns = new HashSet<int>();

            var fields = new[]
            {
                new { Name = "Фамилия", Control = numSurname },
                new { Name = "Имя", Control = numName },
                new { Name = "Группа", Control = numGroup },
                new { Name = "Предпоследняя ЭС", Control = numPrev },
                new { Name = "Последняя ЭС", Control = numLast },
                new { Name = "Средний балл", Control = numMean }
            };

            foreach (var field in fields)
            {
                int value = (int)field.Control.Value;

                if (value < 1)
                {
                    MessageBox.Show($"Поле \"{field.Name}\" не заполнено.\nВсе обязательные поля должны быть заполнены.",
                                   "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    field.Control.Focus();
                    return false;
                }

                if (!columns.Add(value))
                {
                    MessageBox.Show($"Столбец {value} указан несколько раз.\nНомера столбцов не должны повторяться.",
                                   "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            return true;
        }

        private void LockControls(bool locked)
        {
            isProcessing = locked;
            numSurname.Enabled = !locked;
            numName.Enabled = !locked;
            numGroup.Enabled = !locked;
            numPrev.Enabled = !locked;
            numLast.Enabled = !locked;
            numMean.Enabled = !locked;
            numDelay.Enabled = !locked;
            btnBrowse.Enabled = !locked;
            numRoom.Enabled = !locked;
            txtFilePath.AllowDrop = !locked;
        }

        private void CreateGradeTab(string title, Color color)
        {
            TabPage tab = new TabPage(title);

            ListView lv = new ListView
            {
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Location = new System.Drawing.Point(5, 5),
                Size = new System.Drawing.Size(780, 230),
                BackColor = Color.White
            };

            lv.Columns.Add("Фамилия", 180);
            lv.Columns.Add("Имя", 150);
            lv.Columns.Add("Группа", 80);
            lv.Columns.Add("Комната", 80);
            lv.Columns.Add("Балл", 80);

            tab.Controls.Add(lv);
            tabControl.TabPages.Add(tab);

            gradeLists[title] = lv;
        }

        private string GetGradeKey(double meanMark)
        {
            if (meanMark >= 5.0) return "5.0";
            if (meanMark >= 4.5) return "4.5-4.99";
            if (meanMark >= 4.0) return "4.0-4.49";
            if (meanMark >= 3.5) return "3.5-3.99";
            if (meanMark >= 3.0) return "3.0-3.49";
            if (meanMark >= 2.5) return "2.5-2.99";
            return "< 2.5";
        }

        private void BtnClearLog_Click(object sender, EventArgs e)
        {
            logBox.Clear();
            Log("Лог очищен.");
        }
    }
}