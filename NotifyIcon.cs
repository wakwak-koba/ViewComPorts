using System.Linq;

namespace ViewComPorts
{
    public partial class NotifyIcon : System.ComponentModel.Component
    {
        System.Windows.Forms.NotifyIcon icon = new System.Windows.Forms.NotifyIcon();
        Forms.Nickname frmSettings;

        public NotifyIcon()
        {
            InitializeComponent();

            icon.Visible = true;
            icon.Icon = Properties.Resources.ViewComPorts;
            icon.Text = "ViewComPorts";
            icon.ContextMenuStrip = new System.Windows.Forms.ContextMenuStrip();
            icon.ContextMenuStrip.Opening += (sender, e) =>
            {
                Refresh(sender as System.Windows.Forms.ContextMenuStrip);
                e.Cancel = false;
            };
            icon.MouseUp += (sender, e) =>
            {
                if (e.Button == System.Windows.Forms.MouseButtons.Left)
                {
                    var mi = sender.GetType().GetMethod("ShowContextMenu", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                    mi?.Invoke(sender, null);
                }
            };
            Refresh(icon.ContextMenuStrip);
        }

        private void Refresh(System.Windows.Forms.ContextMenuStrip menuStrip)
        {
            menuStrip.Items.Clear();

            // ComPorts
            var Nicknames = new General.DictionaryEx<string, string>(Properties.Settings.Default.Nickname);
            menuStrip.Items.AddRange(new General.ComPorts().Select(m => new System.Windows.Forms.ToolStripLabel(Nicknames.ContainsKey(m.DeviceID) ? Nicknames[m.DeviceID] + " (COM" + m.Port.ToString() + ")" : m.Name)).ToArray());

            // Settings
            menuStrip.Items.Add(new System.Windows.Forms.ToolStripSeparator());
            var btnClose = new System.Windows.Forms.ToolStripMenuItem("Nickname Settings");
            btnClose.Click += (sender, e) => {
                if (frmSettings != null)
                    frmSettings.Dispose();
                frmSettings = new Forms.Nickname(new General.ComPorts());
                frmSettings.ShowDialog();
            };
            menuStrip.Items.Add(btnClose);

            // Exit
            btnClose = new System.Windows.Forms.ToolStripMenuItem("Exit Application");
            btnClose.Click += (sender, e) => {
                icon.Dispose();
                System.Windows.Forms.Application.Exit();
            };
            menuStrip.Items.Add(btnClose);
        }
    }
}
