using System;
using System.Windows.Forms;

namespace AbbreviationWordAddin
{
    public partial class ProgressForm : Form
    {
        private ProgressBar progressBar;
        private Label statusLabel;

        public ProgressForm()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            // Form settings
            this.Text = "Processing...";
            this.Size = new System.Drawing.Size(400, 150);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ControlBox = false;

            // Progress bar
            progressBar = new ProgressBar();
            progressBar.Location = new System.Drawing.Point(20, 20);
            progressBar.Size = new System.Drawing.Size(345, 30);
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Value = 0;

            // Status label
            statusLabel = new Label();
            statusLabel.Location = new System.Drawing.Point(20, 60);
            statusLabel.Size = new System.Drawing.Size(345, 30);
            statusLabel.Text = "Processing...";
            statusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;

            // Add controls to form
            this.Controls.Add(progressBar);
            this.Controls.Add(statusLabel);
        }

        public void UpdateProgress(int percentage, string status = null)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UpdateProgress(percentage, status)));
                return;
            }

            progressBar.Value = percentage;
            if (!string.IsNullOrEmpty(status))
            {
                statusLabel.Text = status;
            }
        }
    }
}
