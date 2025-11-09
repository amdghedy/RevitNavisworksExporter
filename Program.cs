using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace RevitNavisworksExporterEXE
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    public class MainForm : Form
    {
        private TextBox txtFolderPath;
        private Button btnBrowse;
        private Button btnExport;
        private CheckBox chkIncludeSubfolders;
        private CheckBox chkShowRevit;
        private TextBox txtLog;
        private Label lblStatus;
        private ProgressBar progressBar;
        
        private string logFilePath;
        private int totalFiles = 0;
        private int processedFiles = 0;
        private int successCount = 0;
        private int failCount = 0;

        public MainForm()
        {
            InitializeComponents();
            CheckRevitInstallation();
        }

        private void InitializeComponents()
        {
            this.Text = "Revit to Navisworks Batch Exporter";
            this.Size = new System.Drawing.Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new System.Drawing.Size(800, 600);

            // Folder selection
            Label lblFolder = new Label
            {
                Text = "Select Folder:",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(100, 20)
            };

            txtFolderPath = new TextBox
            {
                Location = new System.Drawing.Point(120, 18),
                Size = new System.Drawing.Size(500, 25),
                ReadOnly = true
            };

            btnBrowse = new Button
            {
                Text = "Browse...",
                Location = new System.Drawing.Point(630, 16),
                Size = new System.Drawing.Size(100, 28)
            };
            btnBrowse.Click += BtnBrowse_Click;

            // Options
            chkIncludeSubfolders = new CheckBox
            {
                Text = "Include Subfolders",
                Location = new System.Drawing.Point(120, 55),
                Size = new System.Drawing.Size(200, 20),
                Checked = true
            };

            chkShowRevit = new CheckBox
            {
                Text = "Show Revit Window (visible processing)",
                Location = new System.Drawing.Point(120, 80),
                Size = new System.Drawing.Size(300, 20),
                Checked = true
            };

            // Export button
            btnExport = new Button
            {
                Text = "Start Export",
                Location = new System.Drawing.Point(120, 115),
                Size = new System.Drawing.Size(150, 35),
                Enabled = false,
                BackColor = System.Drawing.Color.FromArgb(0, 120, 215),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnExport.Click += BtnExport_Click;

            // Status
            lblStatus = new Label
            {
                Text = "Ready. Please select a folder containing Revit files.",
                Location = new System.Drawing.Point(20, 165),
                Size = new System.Drawing.Size(750, 20)
            };

            // Progress bar
            progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(20, 190),
                Size = new System.Drawing.Size(750, 25),
                Visible = false
            };

            // Log text box
            Label lblLog = new Label
            {
                Text = "Processing Log:",
                Location = new System.Drawing.Point(20, 225),
                Size = new System.Drawing.Size(100, 20)
            };

            txtLog = new TextBox
            {
                Location = new System.Drawing.Point(20, 250),
                Size = new System.Drawing.Size(750, 280),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new System.Drawing.Font("Consolas", 9)
            };

            // Add controls
            this.Controls.AddRange(new Control[] {
                lblFolder, txtFolderPath, btnBrowse,
                chkIncludeSubfolders, chkShowRevit,
                btnExport, lblStatus, progressBar,
                lblLog, txtLog
            });
        }

        private void CheckRevitInstallation()
        {
            string revitPath = @"C:\Program Files\Autodesk\Revit 2020\Revit.exe";
            if (!File.Exists(revitPath))
            {
                MessageBox.Show(
                    "Revit 2020 not found at standard location.\n" +
                    "Please ensure Revit 2020 is installed.",
                    "Revit Not Found",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
            }
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select folder containing Revit files";
                folderDialog.ShowNewFolderButton = false;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFolderPath.Text = folderDialog.SelectedPath;
                    btnExport.Enabled = true;
                    
                    // Count files
                    SearchOption searchOption = chkIncludeSubfolders.Checked 
                        ? SearchOption.AllDirectories 
                        : SearchOption.TopDirectoryOnly;
                    
                    string[] files = Directory.GetFiles(
                        folderDialog.SelectedPath, "*.rvt", searchOption);
                    
                    lblStatus.Text = $"Found {files.Length} Revit file(s). Click 'Start Export' to begin.";
                }
            }
        }

        private async void BtnExport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFolderPath.Text))
            {
                MessageBox.Show("Please select a folder first.", "No Folder Selected",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Disable controls during processing
            btnExport.Enabled = false;
            btnBrowse.Enabled = false;
            chkIncludeSubfolders.Enabled = false;
            chkShowRevit.Enabled = false;
            progressBar.Visible = true;
            txtLog.Clear();

            try
            {
                await System.Threading.Tasks.Task.Run(() => ProcessFiles());
                
                MessageBox.Show(
                    $"Export Complete!\n\n" +
                    $"Total Files: {totalFiles}\n" +
                    $"Successful: {successCount}\n" +
                    $"Failed: {failCount}\n\n" +
                    $"Log saved to: {Path.GetFileName(logFilePath)}",
                    "Export Complete",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"An error occurred:\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                // Re-enable controls
                btnExport.Enabled = true;
                btnBrowse.Enabled = true;
                chkIncludeSubfolders.Enabled = true;
                chkShowRevit.Enabled = true;
                progressBar.Visible = false;
                UpdateStatus("Ready. Select another folder or close the application.");
            }
        }

        private void ProcessFiles()
        {
            string folderPath = txtFolderPath.Text;
            SearchOption searchOption = chkIncludeSubfolders.Checked 
                ? SearchOption.AllDirectories 
                : SearchOption.TopDirectoryOnly;

            // Initialize log
            logFilePath = Path.Combine(folderPath, 
                $"NavisworksExport_Log_{DateTime.Now:yyyyMMdd_HHmmss}.txt");

            List<string> logMessages = new List<string>();
            logMessages.Add("=== Batch Navisworks Export Started ===");
            logMessages.Add($"Date: {DateTime.Now}");
            logMessages.Add($"Source Folder: {folderPath}");
            logMessages.Add($"Include Subfolders: {chkIncludeSubfolders.Checked}");
            logMessages.Add("");

            AppendLog(string.Join("\n", logMessages));

            // Find all Revit files
            string[] revitFiles = Directory.GetFiles(folderPath, "*.rvt", searchOption);
            totalFiles = revitFiles.Length;
            processedFiles = 0;
            successCount = 0;
            failCount = 0;

            UpdateProgress(0, totalFiles);
            AppendLog($"Found {totalFiles} Revit file(s) to process\n");

            // Process each file
            foreach (string filePath in revitFiles)
            {
                processedFiles++;
                string fileName = Path.GetFileName(filePath);
                
                UpdateStatus($"Processing {processedFiles}/{totalFiles}: {fileName}");
                UpdateProgress(processedFiles, totalFiles);
                AppendLog($"[{processedFiles}/{totalFiles}] Processing: {fileName}");

                try
                {
                    ProcessSingleFile(filePath, logMessages);
                    successCount++;
                    AppendLog("✓ SUCCESS: Exported to NWC\n");
                }
                catch (Exception ex)
                {
                    failCount++;
                    AppendLog($"✗ FAILED: {ex.Message}\n");
                    logMessages.Add($"✗ FAILED: {ex.Message}");
                }

                logMessages.Add("");
            }

            // Final summary
            logMessages.Add("=== Export Complete ===");
            logMessages.Add($"Total Files: {totalFiles}");
            logMessages.Add($"Successful: {successCount}");
            logMessages.Add($"Failed: {failCount}");
            logMessages.Add($"Completion Time: {DateTime.Now}");

            AppendLog("\n=== Export Complete ===");
            AppendLog($"Total Files: {totalFiles}");
            AppendLog($"Successful: {successCount}");
            AppendLog($"Failed: {failCount}");

            // Save log file
            File.WriteAllLines(logFilePath, logMessages);
        }

        private void ProcessSingleFile(string filePath, List<string> logMessages)
        {
            // Create journal file for Revit automation
            string journalPath = Path.Combine(
                Path.GetTempPath(), 
                $"NavisworksExport_{Guid.NewGuid()}.txt"
            );

            string nwcPath = Path.ChangeExtension(filePath, ".nwc");
            
            // Simple check if file might be workshared (just checks file size and name patterns)
            bool possiblyWorkshared = CheckIfPossiblyWorkshared(filePath);
            
            if (possiblyWorkshared)
            {
                AppendLog("  → File may be workshared, will attempt to open as local");
                logMessages.Add("  → File may be workshared, will attempt to open as local");
            }

            // Create Revit journal file for automation
            CreateJournalFile(journalPath, filePath, nwcPath, possiblyWorkshared);

            // Launch Revit with journal
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = @"C:\Program Files\Autodesk\Revit 2020\Revit.exe",
                Arguments = journalPath,
                UseShellExecute = false,
                CreateNoWindow = !chkShowRevit.Checked,
                WindowStyle = chkShowRevit.Checked ? ProcessWindowStyle.Normal : ProcessWindowStyle.Hidden
            };

            AppendLog($"  → Exporting to: {Path.GetFileName(nwcPath)}");
            logMessages.Add($"  → Exporting to: {Path.GetFileName(nwcPath)}");

            using (Process process = Process.Start(startInfo))
            {
                // Wait for Revit to complete (timeout after 10 minutes per file)
                if (!process.WaitForExit(600000)) // 10 minutes
                {
                    process.Kill();
                    throw new Exception("Export timed out after 10 minutes");
                }

                // Check if export was successful
                if (!File.Exists(nwcPath))
                {
                    throw new Exception("NWC file was not created");
                }

                // Verify file size
                FileInfo nwcInfo = new FileInfo(nwcPath);
                if (nwcInfo.Length < 1000) // Less than 1KB suggests failure
                {
                    throw new Exception("NWC file is too small, export may have failed");
                }

                AppendLog($"  → Export completed ({nwcInfo.Length / 1024} KB)");
                logMessages.Add($"  → Export completed ({nwcInfo.Length / 1024} KB)");
            }

            // Clean up journal file
            try { File.Delete(journalPath); } catch { }
        }

        private bool CheckIfPossiblyWorkshared(string filePath)
        {
            // Simple heuristic: check if filename contains common workshared patterns
            string fileName = Path.GetFileName(filePath).ToLower();
            
            // Common patterns in workshared files
            if (fileName.Contains("_central") || 
                fileName.Contains("central_") ||
                fileName.Contains("_detached") ||
                fileName.Contains("backup"))
            {
                return true;
            }

            // Check file size (workshared files are typically larger)
            FileInfo fileInfo = new FileInfo(filePath);
            if (fileInfo.Length > 50 * 1024 * 1024) // > 50 MB
            {
                return true;
            }

            return false;
        }

        private void CreateJournalFile(string journalPath, string revitFile, 
            string nwcFile, bool possiblyWorkshared)
        {
            List<string> journalLines = new List<string>
            {
                "' Revit Batch Export Journal",
                "Dim Jrn",
                "Set Jrn = CrsJournalScript",
                ""
            };

            // Open file
            journalLines.Add($"Jrn.Command \"Ribbon\" , \"Open an existing project , ID_REVIT_FILE_OPEN\"");
            journalLines.Add($"Jrn.Data \"File Name\" , \"IDOK\", \"{revitFile}\"");
            
            if (possiblyWorkshared)
            {
                // Try to handle detach dialog if it appears
                journalLines.Add($"Jrn.Data \"WorksetConfig\" , \"All\", 1");
                journalLines.Add($"Jrn.PushButton \"Detach Model\" , \"Detach and preserve worksets\", \"IDOK\"");
            }

            // Suppress warnings
            journalLines.Add($"Jrn.Data \"Suppress Warnings\" , \"True\"");

            // Export to Navisworks
            journalLines.Add($"Jrn.Command \"Ribbon\" , \"Export to Navisworks , ID_EXPORT_NAVISWORKS\"");
            journalLines.Add($"Jrn.Data \"Export File Name\" , \"{nwcFile}\"");
            journalLines.Add($"Jrn.PushButton \"Export\" , \"OK\", \"IDOK\"");
            
            // Close without saving
            journalLines.Add($"Jrn.Command \"Internal\" , \"Close the active project , ID_REVIT_FILE_CLOSE\"");
            journalLines.Add($"Jrn.PushButton \"Save Changes\" , \"No\", \"IDNO\"");
            
            // Exit Revit
            journalLines.Add($"Jrn.Command \"SystemMenu\" , \"Quit the application; prompts to save projects , ID_APP_EXIT\"");

            File.WriteAllLines(journalPath, journalLines);
        }

        private void UpdateStatus(string status)
        {
            if (lblStatus.InvokeRequired)
            {
                lblStatus.Invoke(new Action(() => lblStatus.Text = status));
            }
            else
            {
                lblStatus.Text = status;
            }
        }

        private void UpdateProgress(int current, int total)
        {
            if (progressBar.InvokeRequired)
            {
                progressBar.Invoke(new Action(() =>
                {
                    progressBar.Maximum = total;
                    progressBar.Value = current;
                }));
            }
            else
            {
                progressBar.Maximum = total;
                progressBar.Value = current;
            }
        }

        private void AppendLog(string message)
        {
            if (txtLog.InvokeRequired)
            {
                txtLog.Invoke(new Action(() =>
                {
                    txtLog.AppendText(message + "\r\n");
                    txtLog.SelectionStart = txtLog.Text.Length;
                    txtLog.ScrollToCaret();
                }));
            }
            else
            {
                txtLog.AppendText(message + "\r\n");
                txtLog.SelectionStart = txtLog.Text.Length;
                txtLog.ScrollToCaret();
            }
        }
    }
}