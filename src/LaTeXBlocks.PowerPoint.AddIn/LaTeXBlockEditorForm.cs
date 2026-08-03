using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LaTeXBlocks.PowerPoint
{
    internal sealed class LaTeXBlockEditorForm : Form
    {
        private readonly PowerPointBlockService service;
        private readonly TextBox sourceBox;
        private readonly TrackBar widthSlider;
        private readonly NumericUpDown widthBox;
        private readonly NumericUpDown fontSizeBox;
        private readonly NumericUpDown lineSpacingBox;
        private readonly NumericUpDown paddingBox;
        private readonly NumericUpDown borderThicknessBox;
        private readonly ComboBox profileBox;
        private readonly ComboBox verticalAlignmentBox;
        private readonly Button textColorButton;
        private readonly Button backgroundColorButton;
        private readonly Button borderColorButton;
        private readonly CheckBox backgroundFillCheck;
        private readonly CheckBox borderCheck;
        private readonly WebBrowser previewBrowser;
        private readonly Button previewButton;
        private readonly Button acceptButton;
        private readonly Label statusLabel;
        private readonly Timer previewTimer;
        private readonly Timer previewNavigationTimer;
        private readonly ToolTip statusToolTip = new ToolTip();
        private readonly Action<string> profileChanged;
        private readonly List<string> previewFiles = new List<string>();
        private bool synchronizingWidth;
        private bool synchronizingProfile;
        private string lastAcceptedProfile;
        private int editVersion;
        private int activeRenders;
        private int renderedVersion = -1;
        private LaTeXBlockRender currentRender;
        private LaTeXBlockRender acceptedRender;
        private string acceptedSource;
        private double acceptedWidthPt;
        private LaTeXBlockStyle acceptedStyle;
        private Color textColor;
        private Color backgroundColor;
        private Color borderColor;
        private readonly double? previewFrameHeightPt;
        private readonly double? previewFrameWidthPt;
        private LaTeXBlockRender pendingPreview;
        private int pendingPreviewVersion = -1;
        private Uri pendingPreviewUri;
        private int previewNavigationAttempts;

        // PowerPoint installs an OLE message filter while a Ribbon callback is
        // unwinding. The hosted WebBrowser is another COM control, so Navigate2
        // can temporarily receive RPC_E_SERVERCALL_RETRYLATER even though the TeX
        // render has already succeeded. Defer and retry that display-only call on
        // the normal WinForms message loop; never report it as a TeX failure.
        private const int RpcEServerCallRetryLater = unchecked((int)0x8001010A);
        private const int RpcECallRejected = unchecked((int)0x80010001);

        internal LaTeXBlockEditorForm(PowerPointBlockService service, string source,
            double widthPt, double fontSizePt, string profile,
            Action<string> profileChanged, bool editing, LaTeXBlockStyle style = null,
            double? outerHeightPt = null, double? outerWidthPt = null)
        {
            this.service = service ?? throw new ArgumentNullException(nameof(service));
            this.profileChanged = profileChanged ?? throw new ArgumentNullException(nameof(profileChanged));
            style = style ?? LaTeXBlockStyle.Default;
            previewFrameHeightPt = outerHeightPt;
            previewFrameWidthPt = outerWidthPt;
            textColor = style.TextColor;
            backgroundColor = style.BackgroundColor;
            borderColor = style.BorderColor;
            widthPt = BlockLayoutWidthPolicy.Clamp(widthPt);
            Text = editing ? "Edit LaTeX Block" : "Insert LaTeX Block";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(840, 600);
            Size = new Size(980, 720);
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
                Minimum = (int)(BlockLayoutWidthPolicy.MinimumPt / BlockLayoutWidthPolicy.StepPt),
                Maximum = (int)(BlockLayoutWidthPolicy.MaximumPt / BlockLayoutWidthPolicy.StepPt),
                SmallChange = 1,
                LargeChange = (int)(10.0 / BlockLayoutWidthPolicy.StepPt),
                TickFrequency = (int)(10.0 / BlockLayoutWidthPolicy.StepPt),
                Value = (int)Math.Round(widthPt / BlockLayoutWidthPolicy.StepPt),
                Width = 140,
                AutoSize = false,
                Height = 28,
                Margin = new Padding(0, 0, 0, 0)
            };
            widthBox = new NumericUpDown
            {
                Minimum = (decimal)BlockLayoutWidthPolicy.MinimumPt,
                Maximum = (decimal)BlockLayoutWidthPolicy.MaximumPt,
                DecimalPlaces = 1,
                Increment = (decimal)BlockLayoutWidthPolicy.StepPt,
                Value = (decimal)widthPt,
                Width = 66
            };
            fontSizeBox = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 200,
                DecimalPlaces = 1,
                Increment = 0.5M,
                Value = (decimal)Math.Max(1, Math.Min(200, fontSizePt)),
                Width = 72
            };
            lineSpacingBox = new NumericUpDown
            {
                Minimum = (decimal)LaTeXBlockStyle.MinimumLineSpacing,
                Maximum = (decimal)LaTeXBlockStyle.MaximumLineSpacing,
                DecimalPlaces = 2,
                Increment = 0.05M,
                Value = (decimal)style.LineSpacing,
                Width = 58
            };
            paddingBox = new NumericUpDown
            {
                Minimum = 0,
                Maximum = (decimal)LaTeXBlockStyle.MaximumPaddingPt,
                DecimalPlaces = 1,
                Increment = 0.5M,
                Value = (decimal)style.PaddingPt,
                Width = 58
            };
            borderThicknessBox = new NumericUpDown
            {
                Minimum = 0,
                Maximum = (decimal)LaTeXBlockStyle.MaximumBorderThicknessPt,
                DecimalPlaces = 1,
                Increment = 0.25M,
                Value = (decimal)style.BorderThicknessPt,
                Width = 54,
                Enabled = style.HasBorder
            };
            profileBox = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 145
            };
            foreach (var availableProfile in service.Profiles)
                profileBox.Items.Add(availableProfile);
            if (profileBox.Items.Count == 0)
                throw new InvalidOperationException("StemTeX has no usable profiles.");
            profileBox.SelectedIndex = 0;
            for (var index = 0; index < profileBox.Items.Count; index++)
                if (string.Equals((string)profileBox.Items[index], profile,
                    StringComparison.OrdinalIgnoreCase))
                    profileBox.SelectedIndex = index;
            lastAcceptedProfile = Profile;

            verticalAlignmentBox = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 76
            };
            verticalAlignmentBox.Items.AddRange(new object[] { "Top", "Middle", "Bottom" });
            verticalAlignmentBox.SelectedIndex = VerticalAlignmentToIndex(style.VerticalAlignment);
            statusToolTip.SetToolTip(verticalAlignmentBox,
                "Controls the SVG frame only when the block has a fixed height. " +
                "An auto-height block fits its content.");
            textColorButton = CreateColorButton(textColor);
            backgroundColorButton = CreateColorButton(backgroundColor);
            borderColorButton = CreateColorButton(borderColor);
            backgroundFillCheck = new CheckBox
            {
                Text = "Fill",
                AutoSize = true,
                Checked = style.HasBackgroundFill,
                Margin = new Padding(12, 4, 2, 0)
            };
            borderCheck = new CheckBox
            {
                Text = "Border",
                AutoSize = true,
                Checked = style.HasBorder,
                Margin = new Padding(12, 4, 2, 0)
            };
            backgroundColorButton.Enabled = style.HasBackgroundFill;
            borderColorButton.Enabled = style.HasBorder;

            previewBrowser = new WebBrowser
            {
                Dock = DockStyle.Fill,
                AllowNavigation = true,
                WebBrowserShortcutsEnabled = false
            };
            previewButton = new Button { Text = "Preview", AutoSize = true };
            acceptButton = new Button
            {
                Text = editing ? "Update" : "Insert",
                AutoSize = true,
                Enabled = false
            };
            var cancelButton = new Button
            {
                Text = "Cancel",
                AutoSize = true,
                DialogResult = DialogResult.Cancel
            };
            statusLabel = new Label { Text = "Ready", AutoSize = true, Anchor = AnchorStyles.Left };
            previewTimer = new Timer { Interval = 250 };
            previewNavigationTimer = new Timer { Interval = 100 };

            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 230
            };
            split.Panel1.Padding = new Padding(12, 8, 12, 8);
            split.Panel1.Controls.Add(sourceBox);
            split.Panel2.Padding = new Padding(12, 8, 12, 8);
            split.Panel2.Controls.Add(previewBrowser);

            var top = new Panel
            {
                Dock = DockStyle.Top,
                Height = 100,
                Padding = new Padding(12, 6, 12, 0)
            };
            var primaryRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 28,
                WrapContents = false
            };
            primaryRow.Controls.Add(new Label
            {
                Text = "Global profile:",
                AutoSize = true,
                Margin = new Padding(0, 4, 8, 0)
            });
            primaryRow.Controls.Add(profileBox);
            primaryRow.Controls.Add(new Label
            {
                Text = "TeX font size (pt):",
                AutoSize = true,
                Margin = new Padding(12, 4, 8, 0)
            });
            primaryRow.Controls.Add(fontSizeBox);
            primaryRow.Controls.Add(previewButton);
            var widthRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 28,
                WrapContents = false
            };
            widthRow.Controls.Add(new Label
            {
                Text = "Typesetting width:",
                AutoSize = true,
                Margin = new Padding(0, 4, 6, 0)
            });
            widthRow.Controls.Add(widthSlider);
            widthRow.Controls.Add(widthBox);
            widthRow.Controls.Add(new Label
            {
                Text = "pt",
                AutoSize = true,
                Margin = new Padding(3, 4, 0, 0)
            });
            var styleRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 28,
                WrapContents = false
            };
            styleRow.Controls.Add(new Label
            {
                Text = "Line spacing:",
                AutoSize = true,
                Margin = new Padding(0, 4, 5, 0)
            });
            styleRow.Controls.Add(lineSpacingBox);
            styleRow.Controls.Add(new Label
            {
                Text = "×",
                AutoSize = true,
                Margin = new Padding(3, 4, 8, 0)
            });
            styleRow.Controls.Add(new Label
            {
                Text = "Padding:",
                AutoSize = true,
                Margin = new Padding(0, 4, 5, 0)
            });
            styleRow.Controls.Add(paddingBox);
            styleRow.Controls.Add(new Label
            {
                Text = "pt",
                AutoSize = true,
                Margin = new Padding(3, 4, 10, 0)
            });
            styleRow.Controls.Add(new Label
            {
                Text = "Vertical:",
                AutoSize = true,
                Margin = new Padding(0, 4, 5, 0)
            });
            styleRow.Controls.Add(verticalAlignmentBox);
            styleRow.Controls.Add(new Label
            {
                Text = "Text:",
                AutoSize = true,
                Margin = new Padding(12, 4, 4, 0)
            });
            styleRow.Controls.Add(textColorButton);
            styleRow.Controls.Add(backgroundFillCheck);
            styleRow.Controls.Add(backgroundColorButton);
            styleRow.Controls.Add(borderCheck);
            styleRow.Controls.Add(borderColorButton);
            styleRow.Controls.Add(borderThicknessBox);
            styleRow.Controls.Add(new Label
            {
                Text = "pt",
                AutoSize = true,
                Margin = new Padding(3, 4, 0, 0)
            });
            top.Controls.Add(styleRow);
            top.Controls.Add(widthRow);
            top.Controls.Add(primaryRow);

            var bottom = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 46,
                Padding = new Padding(12, 8, 12, 0),
                FlowDirection = FlowDirection.RightToLeft
            };
            bottom.Controls.Add(cancelButton);
            bottom.Controls.Add(acceptButton);
            bottom.Controls.Add(statusLabel);

            Controls.Add(split);
            Controls.Add(top);
            Controls.Add(bottom);
            AcceptButton = acceptButton;
            CancelButton = cancelButton;

            sourceBox.TextChanged += (sender, args) => QueueLivePreview();
            widthSlider.ValueChanged += (sender, args) =>
            {
                if (synchronizingWidth) return;
                synchronizingWidth = true;
                try
                {
                    widthBox.Value = (decimal)(widthSlider.Value *
                        BlockLayoutWidthPolicy.StepPt);
                }
                finally { synchronizingWidth = false; }
                QueueLivePreview();
            };
            widthBox.ValueChanged += (sender, args) =>
            {
                if (synchronizingWidth) return;
                synchronizingWidth = true;
                try
                {
                    widthSlider.Value = Math.Max(widthSlider.Minimum,
                        Math.Min(widthSlider.Maximum, (int)Math.Round(
                            (double)widthBox.Value / BlockLayoutWidthPolicy.StepPt)));
                }
                finally { synchronizingWidth = false; }
                QueueLivePreview();
            };
            fontSizeBox.ValueChanged += (sender, args) => QueueLivePreview();
            lineSpacingBox.ValueChanged += (sender, args) => QueueLivePreview();
            paddingBox.ValueChanged += (sender, args) => QueueLivePreview();
            verticalAlignmentBox.SelectedIndexChanged += (sender, args) => QueueLivePreview();
            textColorButton.Click += (sender, args) => ChooseColor(textColorButton,
                value => textColor = value);
            backgroundFillCheck.CheckedChanged += (sender, args) =>
            {
                backgroundColorButton.Enabled = backgroundFillCheck.Checked;
                QueueLivePreview();
            };
            backgroundColorButton.Click += (sender, args) => ChooseColor(backgroundColorButton,
                value => backgroundColor = value);
            borderCheck.CheckedChanged += (sender, args) =>
            {
                borderColorButton.Enabled = borderCheck.Checked;
                borderThicknessBox.Enabled = borderCheck.Checked;
                QueueLivePreview();
            };
            borderColorButton.Click += (sender, args) => ChooseColor(borderColorButton,
                value => borderColor = value);
            borderThicknessBox.ValueChanged += (sender, args) => QueueLivePreview();
            profileBox.SelectedIndexChanged += (sender, args) =>
            {
                if (synchronizingProfile) return;
                var attemptedProfile = Profile;
                try
                {
                    profileChanged(attemptedProfile);
                    lastAcceptedProfile = attemptedProfile;
                    QueueLivePreview();
                }
                catch (Exception exception)
                {
                    // A global-profile change may fail during host shutdown or when
                    // its persistence store is unavailable. Keep the editor and the
                    // add-in on the previously committed profile rather than letting
                    // the WinForms event boundary surface an unhandled exception.
                    synchronizingProfile = true;
                    try
                    {
                        for (var index = 0; index < profileBox.Items.Count; index++)
                            if (string.Equals((string)profileBox.Items[index], lastAcceptedProfile,
                                StringComparison.OrdinalIgnoreCase))
                            {
                                profileBox.SelectedIndex = index;
                                break;
                            }
                    }
                    finally { synchronizingProfile = false; }
                    var message = exception.GetBaseException().Message ?? exception.Message;
                    statusLabel.Text = "Profile change failed: " + message;
                    statusToolTip.SetToolTip(statusLabel, exception.ToString());
                    MessageBox.Show(this, message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
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
            previewNavigationTimer.Tick += (sender, args) =>
            {
                previewNavigationTimer.Stop();
                AttemptPreviewNavigation();
            };
            previewBrowser.DocumentCompleted += PreviewBrowserDocumentCompleted;
            acceptButton.Click += (sender, args) =>
            {
                if (string.IsNullOrWhiteSpace(Source))
                {
                    MessageBox.Show(this, "Enter LaTeX source first.", Text,
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (!PreviewIsCurrent)
                {
                    statusLabel.Text = "Wait for the current preview";
                    return;
                }
                acceptedSource = Source;
                acceptedWidthPt = WidthPt;
                acceptedStyle = Style;
                acceptedRender = currentRender;
                DialogResult = DialogResult.OK;
                Close();
            };
            Shown += (sender, args) => QueueLivePreview();
            FormClosed += (sender, args) =>
            {
                previewTimer.Stop();
                previewNavigationTimer.Stop();
                editVersion++;
                try { service.CancelPreview(); }
                catch { }
                foreach (var path in previewFiles)
                    try { File.Delete(path); } catch { }
            };
        }

        internal string Source => sourceBox.Text;
        internal double WidthPt => (double)widthBox.Value;
        internal double FontSizePt => (double)fontSizeBox.Value;
        internal string Profile => (string)profileBox.SelectedItem;
        internal LaTeXBlockStyle Style => new LaTeXBlockStyle((double)lineSpacingBox.Value,
            (double)paddingBox.Value, IndexToVerticalAlignment(verticalAlignmentBox.SelectedIndex),
            textColor, backgroundFillCheck.Checked, backgroundColor,
            borderCheck.Checked ? (double)borderThicknessBox.Value : 0, borderColor);
        internal LaTeXBlockRender CurrentRender => currentRender;
        internal bool PreviewIsCurrent => currentRender != null && renderedVersion == editVersion;
        internal string AcceptedSource => acceptedSource;
        internal double AcceptedWidthPt => acceptedWidthPt;
        internal LaTeXBlockStyle AcceptedStyle => acceptedStyle ?? Style;
        internal LaTeXBlockRender AcceptedRender => acceptedRender;
        internal void SetSourceForTest(string source) { sourceBox.Text = source; }
        internal void SetWidthPtForTest(double widthPt)
        {
            widthBox.Value = (decimal)BlockLayoutWidthPolicy.Clamp(widthPt);
        }

        private static Button CreateColorButton(Color color)
        {
            var button = new Button
            {
                AutoSize = false,
                Width = 68,
                Height = 24,
                Margin = new Padding(0, 2, 0, 0),
                FlatStyle = FlatStyle.Flat
            };
            UpdateColorButton(button, color);
            return button;
        }

        private void ChooseColor(Button button, Action<Color> apply)
        {
            if (button == null || apply == null || IsDisposed) return;
            using (var dialog = new ColorDialog
            {
                Color = button.BackColor,
                FullOpen = true,
                AnyColor = true
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                var color = Color.FromArgb(dialog.Color.R, dialog.Color.G, dialog.Color.B);
                apply(color);
                UpdateColorButton(button, color);
                QueueLivePreview();
            }
        }

        private static void UpdateColorButton(Button button, Color color)
        {
            button.BackColor = color;
            button.ForeColor = color.GetBrightness() < 0.45f ? Color.White : Color.Black;
            button.Text = "#" + color.R.ToString("X2") + color.G.ToString("X2") +
                color.B.ToString("X2");
            button.FlatAppearance.BorderColor = Color.FromArgb(120, 120, 120);
        }

        private static int VerticalAlignmentToIndex(LaTeXBlockVerticalAlignment alignment)
        {
            switch (alignment)
            {
                case LaTeXBlockVerticalAlignment.Middle: return 1;
                case LaTeXBlockVerticalAlignment.Bottom: return 2;
                default: return 0;
            }
        }

        private static LaTeXBlockVerticalAlignment IndexToVerticalAlignment(int index)
        {
            switch (index)
            {
                case 1: return LaTeXBlockVerticalAlignment.Middle;
                case 2: return LaTeXBlockVerticalAlignment.Bottom;
                default: return LaTeXBlockVerticalAlignment.Top;
            }
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
            statusLabel.Text = activeRenders > 0
                ? "Rendering; latest input queued..."
                : "Waiting for input pause...";
            previewTimer.Start();
        }

        private async Task PreviewLatestAsync(bool showErrorDialog)
        {
            if (IsDisposed || string.IsNullOrWhiteSpace(Source)) return;
            activeRenders++;
            var version = editVersion;
            var source = Source;
            var width = WidthPt;
            var fontSize = FontSizePt;
            var profile = Profile;
            var style = Style;
            var phase = "Render";
            SetRendering(true, "Rendering...");
            try
            {
                var render = await service.RenderPreviewAsync(source, width, profile, fontSize,
                    style, previewFrameHeightPt, previewFrameWidthPt);
                phase = "Preview update";
                if (IsDisposed) return;
                if (version == editVersion)
                    SchedulePreviewNavigation(render, version);
                else
                    statusLabel.Text = "Newer preview requested...";
            }
            catch (TaskCanceledException)
            {
                // A newer source/profile request superseded this render.
            }
            catch (Exception exception)
            {
                if (!IsDisposed && version == editVersion)
                {
                    WriteErrorLog(exception, source, profile, version);
                    var message = (exception.GetBaseException().Message ?? exception.Message)
                        .Replace('\r', ' ').Replace('\n', ' ').Trim();
                    var shortMessage = message.Length > 90
                        ? message.Substring(0, 87) + "..."
                        : message;
                    statusLabel.Text = phase + " failed: " + shortMessage;
                    statusToolTip.SetToolTip(statusLabel, exception.ToString());
                    if (showErrorDialog)
                        MessageBox.Show(this, exception.Message, Text,
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void SetRendering(bool rendering, string status)
        {
            previewButton.Enabled = !rendering;
            acceptButton.Enabled = !rendering && currentRender != null &&
                                   renderedVersion == editVersion;
            statusLabel.Text = status;
        }

        private void SchedulePreviewNavigation(LaTeXBlockRender render, int version)
        {
            var token = version + "-" + Guid.NewGuid().ToString("N");
            var htmlPath = Path.Combine(Path.GetDirectoryName(render.SvgPath),
                token + ".preview.html");
            File.WriteAllText(htmlPath, BuildPreviewHtml(render.SvgBytes, token),
                new UTF8Encoding(false));
            previewFiles.Add(htmlPath);
            pendingPreview = render;
            pendingPreviewVersion = version;
            pendingPreviewUri = new Uri(htmlPath);
            previewNavigationAttempts = 0;
            statusLabel.Text = "Updating preview...";
            acceptButton.Enabled = false;
            // Post after the current VSTO/Ribbon callback returns. Calling the
            // WebBrowser COM control directly from that callback is rejected by
            // PowerPoint's message filter on some machines.
            try
            {
                BeginInvoke(new MethodInvoker(AttemptPreviewNavigation));
            }
            catch (InvalidOperationException)
            {
                // The form is closing; FormClosed owns cleanup of this preview.
            }
        }

        private void AttemptPreviewNavigation()
        {
            if (IsDisposed || pendingPreviewUri == null ||
                pendingPreviewVersion != editVersion) return;
            try
            {
                previewBrowser.Navigate(pendingPreviewUri);
            }
            catch (COMException exception) when (IsTransientPreviewNavigationFailure(exception))
            {
                previewNavigationAttempts++;
                // The retry itself is deliberately low-cost: it merely posts one
                // existing local HTML URI to the already-created browser control.
                // Back off a little while PowerPoint finishes its OLE callback.
                previewNavigationTimer.Interval = Math.Min(500,
                    80 + previewNavigationAttempts * 20);
                statusLabel.Text = "PowerPoint is busy; retrying preview...";
                previewNavigationTimer.Start();
            }
            catch (Exception exception)
            {
                HandlePreviewNavigationFailure(exception);
            }
        }

        internal static bool IsTransientPreviewNavigationFailure(COMException exception)
        {
            return exception != null && (exception.ErrorCode == RpcEServerCallRetryLater ||
                exception.ErrorCode == RpcECallRejected);
        }

        private void HandlePreviewNavigationFailure(Exception exception)
        {
            if (IsDisposed || pendingPreviewVersion != editVersion) return;
            WriteErrorLog(exception, Source, Profile, pendingPreviewVersion);
            var message = (exception.GetBaseException().Message ?? exception.Message)
                .Replace('\r', ' ').Replace('\n', ' ').Trim();
            var shortMessage = message.Length > 90 ? message.Substring(0, 87) + "..." : message;
            statusLabel.Text = "Preview display failed: " + shortMessage;
            statusToolTip.SetToolTip(statusLabel, exception.ToString());
        }

        private void PreviewBrowserDocumentCompleted(object sender,
            WebBrowserDocumentCompletedEventArgs args)
        {
            if (IsDisposed || pendingPreviewUri == null || args.Url == null) return;
            if (!string.Equals(args.Url.LocalPath, pendingPreviewUri.LocalPath,
                    StringComparison.OrdinalIgnoreCase)) return;
            if (previewBrowser.ReadyState != WebBrowserReadyState.Complete) return;

            var marker = previewBrowser.Document?.GetElementById("latexblocks-preview-token");
            var expectedToken = Path.GetFileNameWithoutExtension(
                Path.GetFileNameWithoutExtension(pendingPreviewUri.LocalPath));
            if (marker == null || !string.Equals(marker.GetAttribute("data-token"),
                    expectedToken, StringComparison.Ordinal)) return;
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
            var svg = Encoding.UTF8.GetString(svgBytes ??
                throw new ArgumentNullException(nameof(svgBytes)));
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

        private static void WriteErrorLog(Exception exception, string source,
            string profile, int version)
        {
            try
            {
                var directory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "LaTeXBlocks");
                Directory.CreateDirectory(directory);
                File.AppendAllText(Path.Combine(directory, "errors.log"),
                    DateTime.Now.ToString("O") + " PowerPoint preview version=" + version +
                    " profile=" + profile + Environment.NewLine + source + Environment.NewLine +
                    exception + Environment.NewLine + new string('-', 72) + Environment.NewLine,
                    new UTF8Encoding(false));
            }
            catch { }
        }
    }
}
