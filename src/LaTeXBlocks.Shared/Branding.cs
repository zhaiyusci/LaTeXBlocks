using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace LaTeXBlocks.Branding
{
    internal static class BrandAssets
    {
        private const string IconResource = "LaTeXBlocks.Branding.Icon.ico";
        private const string LogoResource = "LaTeXBlocks.Branding.Icon.png";

        internal static Icon LoadIcon()
        {
            using (var stream = OpenResource(IconResource))
            using (var icon = new Icon(stream))
                return (Icon)icon.Clone();
        }

        internal static Image LoadLogo()
        {
            using (var stream = OpenResource(LogoResource))
            using (var image = Image.FromStream(stream))
                return new Bitmap(image);
        }

        internal static void ApplyTo(Form form)
        {
            if (form == null) throw new ArgumentNullException(nameof(form));
            var icon = LoadIcon();
            form.Icon = icon;
            form.Disposed += (sender, args) => icon.Dispose();
        }

        private static Stream OpenResource(string name)
        {
            var stream = typeof(BrandAssets).Assembly.GetManifestResourceStream(name);
            if (stream == null)
                throw new InvalidOperationException("Missing embedded branding resource: " + name);
            return stream;
        }
    }

    internal sealed class AboutForm : Form
    {
        private const string ProjectUrl = "https://github.com/zhaiyusci/LaTeXBlocks";
        private const string SupportUrl = ProjectUrl + "/issues";

        internal AboutForm(string hostName)
        {
            Text = "About LaTeX Blocks";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(560, 270);
            BrandAssets.ApplyTo(this);

            var logo = new PictureBox
            {
                Image = BrandAssets.LoadLogo(),
                SizeMode = PictureBoxSizeMode.Zoom,
                Location = new Point(28, 28),
                Size = new Size(176, 176),
                TabStop = false
            };

            var name = new Label
            {
                AutoSize = true,
                Font = new Font(Font.FontFamily, 20F, FontStyle.Bold),
                ForeColor = Color.FromArgb(9, 43, 112),
                Location = new Point(230, 30),
                Text = "LaTeX Blocks"
            };
            var tagline = new Label
            {
                AutoSize = false,
                Location = new Point(234, 76),
                Size = new Size(294, 45),
                Text = "Editable LaTeX blocks for Microsoft " + hostName + "."
            };
            var version = new Label
            {
                AutoSize = true,
                Location = new Point(234, 126),
                Text = "Version " + GetDisplayVersion()
            };
            var project = CreateLink("Project home", ProjectUrl, new Point(234, 158));
            var support = CreateLink("Report an issue", SupportUrl, new Point(334, 158));
            var copyright = new Label
            {
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Location = new Point(234, 193),
                Text = "Copyright © 2026 Y. Zhai"
            };
            var close = new Button
            {
                DialogResult = DialogResult.OK,
                Location = new Point(441, 226),
                Size = new Size(88, 28),
                Text = "OK"
            };

            Controls.AddRange(new Control[]
            {
                logo, name, tagline, version, project, support, copyright, close
            });
            AcceptButton = close;
            CancelButton = close;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                foreach (Control control in Controls)
                {
                    var picture = control as PictureBox;
                    if (picture != null)
                        picture.Image?.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        private static LinkLabel CreateLink(string text, string url, Point location)
        {
            var link = new LinkLabel { AutoSize = true, Location = location, Text = text };
            link.LinkClicked += (sender, args) => OpenUrl(url);
            return link;
        }

        private static void OpenUrl(string url)
        {
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        }

        private static string GetDisplayVersion()
        {
            var version = typeof(AboutForm).Assembly.GetName().Version;
            if (version == null)
                return "unknown";
            return version.Revision == 0
                ? version.Major + "." + version.Minor + "." + version.Build
                : version.ToString();
        }
    }
}
