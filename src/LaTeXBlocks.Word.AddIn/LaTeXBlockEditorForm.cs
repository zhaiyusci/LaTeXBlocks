using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LaTeXBlocks.Word
{
    internal sealed class LaTeXBlockEditorForm : Form
    {
        private readonly LaTeXBlockService service;
        private readonly TextBox sourceBox;
        private readonly TrackBar widthSlider;
        private readonly NumericUpDown widthBox;
        private readonly Label widthUnitLabel;
        private readonly Label naturalWidthLabel;
        private readonly ComboBox modeBox;
        private readonly ComboBox profileBox;
        private readonly WebBrowser previewBrowser;
        private readonly Button previewButton;
        private readonly Button acceptButton;
        private readonly Label statusLabel;
        private readonly Timer previewTimer;
        private readonly ToolTip statusToolTip = new ToolTip();
        private readonly Action<string> profileChanged;
        private readonly double fontSizePt;
        private readonly bool displayMathStyle;
        private readonly double naturalWidthSentinelPt;
        private bool synchronizingWidth;
        private bool synchronizingProfile;
        private string acceptedProfile;
        private int editVersion;
        private int activeRenders;
        private int renderedVersion = -1;
        private LaTeXBlockRender currentRender;
        private LaTeXBlockRender pendingPreview;
        private int pendingPreviewVersion = -1;
        private Uri pendingPreviewUri;
        private readonly System.Collections.Generic.List<string> previewFiles =
            new System.Collections.Generic.List<string>();

        internal LaTeXBlockEditorForm(LaTeXBlockService service, string source, double widthPt,
            LaTeXBlockLayoutMode mode, string profile,
            Action<string> profileChanged, bool editing, double fontSizePt = 10,
            string windowTitle = null, bool displayMathStyle = false)
        {
            this.service = service ?? throw new ArgumentNullException(nameof(service));
            this.profileChanged = profileChanged ?? throw new ArgumentNullException(nameof(profileChanged));
            this.fontSizePt = fontSizePt;
            this.displayMathStyle = displayMathStyle;
            naturalWidthSentinelPt = widthPt > 0 && !double.IsNaN(widthPt) &&
                !double.IsInfinity(widthPt) ? widthPt : LaTeXBlockWidthPolicy.DefaultWidthPt;
            widthPt = LaTeXBlockWidthPolicy.ClampWidth(widthPt);
            Text = windowTitle ?? (editing ? "Edit LaTeX Block" : "Insert LaTeX Block");
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(760, 560);
            Size = new Size(900, 680);
            Font = SystemFonts.MessageBoxFont;

            sourceBox = new TextBox
            {
                Multiline = true,
                AcceptsReturn = true,
                AcceptsTab = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Font = new Font("Consolas", 11F),
                Text = source ?? string.Empty,
                Dock = DockStyle.Fill
            };
            widthSlider = new TrackBar
            {
                Minimum = (int)(LaTeXBlockWidthPolicy.MinimumWidthPt /
                    LaTeXBlockWidthPolicy.WidthStepPt),
                Maximum = (int)(LaTeXBlockWidthPolicy.MaximumWidthPt /
                    LaTeXBlockWidthPolicy.WidthStepPt),
                SmallChange = 1,
                LargeChange = (int)(10.0 / LaTeXBlockWidthPolicy.WidthStepPt),
                TickFrequency = (int)(10.0 / LaTeXBlockWidthPolicy.WidthStepPt),
                Value = (int)Math.Round(widthPt / LaTeXBlockWidthPolicy.WidthStepPt),
                Width = 140,
                AutoSize = false,
                Height = 28,
                Margin = new Padding(0, 0, 0, 0)
            };
            widthBox = new NumericUpDown
            {
                Minimum = (decimal)LaTeXBlockWidthPolicy.MinimumWidthPt,
                Maximum = (decimal)LaTeXBlockWidthPolicy.MaximumWidthPt,
                DecimalPlaces = 1,
                Increment = (decimal)LaTeXBlockWidthPolicy.WidthStepPt,
                Width = 66,
                Value = (decimal)widthPt
            };
            widthUnitLabel = new Label
            {
                Text = "pt",
                AutoSize = true,
                Margin = new Padding(4, 4, 0, 0)
            };
            naturalWidthLabel = new Label
            {
                Text = "Natural",
                AutoSize = true,
                Margin = new Padding(0, 4, 0, 0)
            };
            modeBox = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 155 };
            if (displayMathStyle)
            {
                modeBox.Items.Add("Display equation (Auto)");
                modeBox.SelectedIndex = 0;
            }
            else
            {
                modeBox.Items.Add("Auto-width formula");
                modeBox.Items.Add("LaTeX block (Fixed)");
                modeBox.SelectedIndex = mode == LaTeXBlockLayoutMode.Auto ? 0 : 1;
            }
            profileBox = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 145 };
            foreach (var availableProfile in service.Profiles) profileBox.Items.Add(availableProfile);
            profileBox.SelectedIndex = 0;
            for (var index = 0; index < profileBox.Items.Count; index++)
                if (string.Equals((string)profileBox.Items[index], profile, StringComparison.OrdinalIgnoreCase))
                    profileBox.SelectedIndex = index;
            acceptedProfile = Profile;

            previewBrowser = new WebBrowser { Dock = DockStyle.Fill, AllowNavigation = true, WebBrowserShortcutsEnabled = false };
            previewButton = new Button { Text = "Preview", AutoSize = true };
            acceptButton = new Button { Text = editing ? "Update" : "Insert", AutoSize = true };
            acceptButton.Enabled = false;
            var cancelButton = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
            statusLabel = new Label { Text = "Ready", AutoSize = true, Anchor = AnchorStyles.Left };
            previewTimer = new Timer { Interval = 250 };

            var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterDistance = 230 };
            split.Panel1.Padding = new Padding(12, 8, 12, 8);
            split.Panel1.Controls.Add(sourceBox);
            split.Panel2.Padding = new Padding(12, 8, 12, 8);
            split.Panel2.Controls.Add(previewBrowser);

            var top = new Panel { Dock = DockStyle.Top, Height = 68, Padding = new Padding(12, 6, 12, 0) };
            var primaryRow = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 28, WrapContents = false };
            primaryRow.Controls.Add(new Label { Text = "Layout:", AutoSize = true, Margin = new Padding(0, 4, 8, 0) });
            primaryRow.Controls.Add(modeBox);
            primaryRow.Controls.Add(new Label { Text = "Global profile:", AutoSize = true, Margin = new Padding(12, 4, 8, 0) });
            primaryRow.Controls.Add(profileBox);
            primaryRow.Controls.Add(previewButton);
            var widthRow = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 28, WrapContents = false };
            widthRow.Controls.Add(new Label { Text = "Typesetting width:", AutoSize = true, Margin = new Padding(0, 4, 6, 0) });
            widthRow.Controls.Add(widthSlider);
            widthRow.Controls.Add(widthBox);
            widthRow.Controls.Add(widthUnitLabel);
            widthRow.Controls.Add(naturalWidthLabel);
            top.Controls.Add(widthRow);
            top.Controls.Add(primaryRow);

            var bottom = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 46, Padding = new Padding(12, 8, 12, 0), FlowDirection = FlowDirection.RightToLeft };
            bottom.Controls.Add(cancelButton);
            bottom.Controls.Add(acceptButton);
            bottom.Controls.Add(statusLabel);

            Controls.Add(split);
            Controls.Add(top);
            Controls.Add(bottom);
            AcceptButton = acceptButton;
            CancelButton = cancelButton;
            UpdateWidthControlVisibility();

            sourceBox.TextChanged += (sender, args) => QueueLivePreview();
            modeBox.SelectedIndexChanged += (sender, args) =>
            {
                UpdateWidthControlVisibility();
                QueueLivePreview();
            };
            profileBox.SelectedIndexChanged += ProfileBox_SelectedIndexChanged;
            widthSlider.ValueChanged += (sender, args) =>
            {
                if (synchronizingWidth) return;
                synchronizingWidth = true;
                try
                {
                    widthBox.Value = (decimal)(widthSlider.Value *
                        LaTeXBlockWidthPolicy.WidthStepPt);
                }
                finally { synchronizingWidth = false; }
                if (Mode == LaTeXBlockLayoutMode.Fixed) QueueLivePreview();
            };
            widthBox.ValueChanged += (sender, args) =>
            {
                if (synchronizingWidth) return;
                synchronizingWidth = true;
                try
                {
                    widthSlider.Value = Math.Max(widthSlider.Minimum,
                        Math.Min(widthSlider.Maximum, (int)Math.Round((double)widthBox.Value /
                            LaTeXBlockWidthPolicy.WidthStepPt)));
                }
                finally { synchronizingWidth = false; }
                if (Mode == LaTeXBlockLayoutMode.Fixed) QueueLivePreview();
            };
            previewButton.Click += async (sender, args) =>
            {
                previewTimer.Stop();
                editVersion++;
                await PreviewLatestAsync(true);
            };
            previewTimer.Tick += async (sender, args) =>
            {
                previewTimer.Stop();
                await PreviewLatestAsync(false);
            };
            previewBrowser.DocumentCompleted += PreviewBrowserDocumentCompleted;
            acceptButton.Click += (sender, args) =>
            {
                if (string.IsNullOrWhiteSpace(Source))
                {
                    MessageBox.Show(this, "Enter LaTeX source first.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                DialogResult = DialogResult.OK;
                Close();
            };
            Shown += (sender, args) => QueueLivePreview();
            FormClosed += (sender, args) =>
            {
                previewTimer.Stop();
                editVersion++;
                try { service.CancelPreview(); }
                catch
                {
                    // The add-in can be shutting down concurrently with this dialog.
                    // Closing an editor must remain immediate even if its backend has
                    // already been disposed.
                }
                foreach (var path in previewFiles)
                    try { File.Delete(path); } catch { }
            };
        }

        internal string Source => sourceBox.Text;
        internal double WidthPt => Mode == LaTeXBlockLayoutMode.Auto
            ? naturalWidthSentinelPt
            : (double)widthBox.Value;
        internal bool WidthIsNatural => Mode == LaTeXBlockLayoutMode.Auto;
        internal LaTeXBlockLayoutMode Mode => modeBox.SelectedIndex == 0 ? LaTeXBlockLayoutMode.Auto : LaTeXBlockLayoutMode.Fixed;
        internal string Profile => (string)profileBox.SelectedItem;
        internal LaTeXBlockRender CurrentRender => currentRender;
        internal bool PreviewIsCurrent => currentRender != null && renderedVersion == editVersion;
        internal void SetSourceForTest(string source) { sourceBox.Text = source; }
        internal void SetWidthPtForTest(double widthPt)
        {
            widthBox.Value = (decimal)LaTeXBlockWidthPolicy.ClampWidth(widthPt);
        }

        private void ProfileBox_SelectedIndexChanged(object sender, EventArgs args)
        {
            if (synchronizingProfile) return;
            var requestedProfile = Profile;
            if (string.Equals(requestedProfile, acceptedProfile, StringComparison.OrdinalIgnoreCase)) return;

            try
            {
                profileChanged(requestedProfile);
                acceptedProfile = requestedProfile;
                QueueLivePreview();
            }
            catch (Exception exception)
            {
                // The combobox has already changed when this event fires. Revert it
                // before reporting the problem so the visible choice, the global
                // Word preference, and subsequent renders keep the same profile.
                synchronizingProfile = true;
                try
                {
                    for (var index = 0; index < profileBox.Items.Count; index++)
                        if (string.Equals((string)profileBox.Items[index], acceptedProfile,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            profileBox.SelectedIndex = index;
                            break;
                        }
                }
                finally { synchronizingProfile = false; }

                var message = (exception.GetBaseException().Message ?? exception.Message)
                    .Replace('\r', ' ').Replace('\n', ' ').Trim();
                var shortMessage = message.Length > 90 ? message.Substring(0, 87) + "..." : message;
                statusLabel.Text = "Profile change failed: " + shortMessage;
                statusToolTip.SetToolTip(statusLabel, exception.ToString());
                MessageBox.Show(this, message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateWidthControlVisibility()
        {
            var fixedWidth = Mode == LaTeXBlockLayoutMode.Fixed;
            widthSlider.Visible = fixedWidth;
            widthBox.Visible = fixedWidth;
            widthUnitLabel.Visible = fixedWidth;
            naturalWidthLabel.Visible = !fixedWidth;
        }

        private void QueueLivePreview()
        {
            if (IsDisposed) return;
            editVersion++;
            acceptButton.Enabled = false;
            previewTimer.Stop();
            if (string.IsNullOrWhiteSpace(Source))
            {
                statusLabel.Text = "Enter LaTeX source";
                return;
            }
            statusLabel.Text = activeRenders > 0 ? "Rendering; latest input queued..." : "Waiting for input pause...";
            previewTimer.Interval = 250;
            previewTimer.Start();
        }

        private async Task PreviewLatestAsync(bool showErrorDialog)
        {
            if (IsDisposed || string.IsNullOrWhiteSpace(Source)) return;
            activeRenders++;
            var version = editVersion;
            var source = Source;
            var width = WidthPt;
            var mode = Mode;
            var profile = Profile;
            var phase = "Render";
            SetRendering(true, "Rendering...");
            try
            {
                var render = await service.RenderPreviewAsync(source, width, mode, profile, fontSizePt,
                    displayMathStyle);
                phase = "Preview update";
                if (IsDisposed) return;
                if (version == editVersion)
                {
                    BeginPreviewNavigation(render, version);
                }
                else
                {
                    statusLabel.Text = "Newer preview requested...";
                }
            }
            catch (TaskCanceledException)
            {
                // The backend discarded this result because a newer request or profile won.
            }
            catch (Exception exception)
            {
                if (!IsDisposed && version == editVersion)
                {
                    WriteErrorLog(exception, source, profile, version);
                    var message = (exception.GetBaseException().Message ?? exception.Message)
                        .Replace('\r', ' ').Replace('\n', ' ').Trim();
                    var shortMessage = message.Length > 90 ? message.Substring(0, 87) + "..." : message;
                    statusLabel.Text = phase + " failed: " + shortMessage;
                    statusToolTip.SetToolTip(statusLabel, exception.ToString());
                    if (showErrorDialog)
                        MessageBox.Show(this, exception.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            finally
            {
                if (!IsDisposed)
                {
                    activeRenders = Math.Max(0, activeRenders - 1);
                    SetRendering(activeRenders > 0, statusLabel.Text);
                }
            }
        }

        private static void WriteErrorLog(Exception exception, string source, string profile, int version)
        {
            try
            {
                var directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "LaTeXBlocks");
                Directory.CreateDirectory(directory);
                File.AppendAllText(Path.Combine(directory, "errors.log"),
                    DateTime.Now.ToString("O") + " preview version=" + version + " profile=" + profile +
                    Environment.NewLine + source + Environment.NewLine + exception + Environment.NewLine +
                    new string('-', 72) + Environment.NewLine, new UTF8Encoding(false));
            }
            catch { }
        }

        private void SetRendering(bool rendering, string status)
        {
            previewButton.Enabled = !rendering;
            acceptButton.Enabled = !rendering && currentRender != null && renderedVersion == editVersion;
            statusLabel.Text = status;
            // Source, layout, and profile remain editable while a render is in flight.
        }

        private void BeginPreviewNavigation(LaTeXBlockRender render, int version)
        {
            var token = version.ToString() + "-" + Guid.NewGuid().ToString("N");
            var htmlPath = Path.Combine(Path.GetDirectoryName(render.SvgPath), token + ".preview.html");
            File.WriteAllText(htmlPath, BuildPreviewHtml(render.SvgBytes, token), new UTF8Encoding(false));
            previewFiles.Add(htmlPath);
            pendingPreview = render;
            pendingPreviewVersion = version;
            pendingPreviewUri = new Uri(htmlPath);
            statusLabel.Text = "Updating preview...";
            acceptButton.Enabled = false;
            previewBrowser.Navigate(pendingPreviewUri);
        }

        private void PreviewBrowserDocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs args)
        {
            if (IsDisposed || pendingPreviewUri == null || args.Url == null) return;
            if (!string.Equals(args.Url.LocalPath, pendingPreviewUri.LocalPath, StringComparison.OrdinalIgnoreCase)) return;
            if (previewBrowser.ReadyState != WebBrowserReadyState.Complete) return;

            var marker = previewBrowser.Document?.GetElementById("latexblocks-preview-token");
            var expectedToken = Path.GetFileNameWithoutExtension(Path.GetFileNameWithoutExtension(pendingPreviewUri.LocalPath));
            if (marker == null || !string.Equals(marker.GetAttribute("data-token"), expectedToken, StringComparison.Ordinal)) return;
            if (pendingPreviewVersion != editVersion) return;

            currentRender = pendingPreview;
            renderedVersion = pendingPreviewVersion;
            pendingPreview = null;
            pendingPreviewVersion = -1;
            pendingPreviewUri = null;
            statusLabel.Text = "Preview is current";
            statusToolTip.SetToolTip(statusLabel, string.Empty);
            acceptButton.Enabled = activeRenders == 0;
        }

        internal static string BuildPreviewHtml(byte[] svgBytes, string token = "test")
        {
            var svg = Encoding.UTF8.GetString(svgBytes ?? throw new ArgumentNullException(nameof(svgBytes)));
            if (svg.StartsWith("<?xml", StringComparison.Ordinal))
            {
                var declarationEnd = svg.IndexOf("?>", StringComparison.Ordinal);
                if (declarationEnd >= 0) svg = svg.Substring(declarationEnd + 2);
            }
            return "<!doctype html><html><head><meta http-equiv='X-UA-Compatible' content='IE=edge'/>" +
                   "<style>html,body{margin:0;padding:4px;background:white;overflow:auto;}svg{display:block;}</style>" +
                   "</head><body><span id='latexblocks-preview-token' data-token='" + token +
                   "' style='display:none'></span>" + svg + "</body></html>";
        }
    }
}
