using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace LaTeXBlocks.Word
{
    /// <summary>
    /// Word's built-in cross-reference picker groups our manual-break display
    /// lines by paragraph. This compact picker instead presents the real,
    /// bookmark-backed LaTeXBlockEq targets one per row.
    /// </summary>
    internal sealed class EquationReferenceForm : Form
    {
        private readonly ListView targets;
        private readonly Button insert;

        internal EquationReferenceForm(IReadOnlyList<LaTeXBlockService.EquationReferenceTarget> references)
        {
            if (references == null) throw new ArgumentNullException(nameof(references));

            Text = "Insert Equation Reference";
            Branding.BrandAssets.ApplyTo(this);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(650, 330);

            var instruction = new Label
            {
                Dock = DockStyle.Top,
                AutoSize = false,
                Height = 42,
                Padding = new Padding(12, 10, 12, 4),
                Text = "Select a numbered LaTeX equation. The inserted reference is a native Word field.",
                TextAlign = ContentAlignment.MiddleLeft
            };

            targets = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                HideSelection = false,
                HeaderStyle = ColumnHeaderStyle.Nonclickable,
                GridLines = true
            };
            targets.Columns.Add("Number", 82, HorizontalAlignment.Center);
            targets.Columns.Add("LaTeX source", 540, HorizontalAlignment.Left);
            foreach (var reference in references)
            {
                var item = new ListViewItem("(" + reference.Number + ")") { Tag = reference };
                item.SubItems.Add(SummarizeSource(reference.Source));
                targets.Items.Add(item);
            }
            targets.SelectedIndexChanged += (_, __) => insert.Enabled = targets.SelectedItems.Count == 1;
            targets.DoubleClick += (_, __) => AcceptSelection();

            insert = new Button { Text = "Insert", DialogResult = DialogResult.None, Enabled = false, AutoSize = true };
            insert.Click += (_, __) => AcceptSelection();
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
            var actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 46,
                Padding = new Padding(8),
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false
            };
            actions.Controls.Add(cancel);
            actions.Controls.Add(insert);

            Controls.Add(targets);
            Controls.Add(actions);
            Controls.Add(instruction);
            AcceptButton = insert;
            CancelButton = cancel;
            if (targets.Items.Count > 0) targets.Items[0].Selected = true;
        }

        internal LaTeXBlockService.EquationReferenceTarget SelectedReference
        {
            get
            {
                return targets.SelectedItems.Count == 1
                    ? targets.SelectedItems[0].Tag as LaTeXBlockService.EquationReferenceTarget
                    : null;
            }
        }

        private void AcceptSelection()
        {
            if (SelectedReference == null) return;
            DialogResult = DialogResult.OK;
            Close();
        }

        private static string SummarizeSource(string source)
        {
            var summary = Regex.Replace(source ?? string.Empty, "\\s+", " ").Trim();
            const int maximumLength = 180;
            return summary.Length <= maximumLength
                ? summary
                : summary.Substring(0, maximumLength - 1) + "…";
        }
    }
}
