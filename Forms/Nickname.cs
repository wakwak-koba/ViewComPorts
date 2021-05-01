using System.Linq;

namespace ViewComPorts.Forms
{
    public partial class Nickname : System.Windows.Forms.Form
    {
        private General.DictionaryEx<string, string> Nicknames = new General.DictionaryEx<string, string>(Properties.Settings.Default.Nickname);

        public Nickname()
        {
            InitializeComponent();
        }

        public Nickname(General.ComPorts ComPorts) : this()
        {
            this.Load += (sender, e) =>
            {
                var dt = new System.Data.DataTable();
                dt.PrimaryKey = new [] { dt.Columns.Add("DeviceID", typeof(string)) };
                dt.Columns.Add("Port", typeof(int)).ReadOnly = true;
                dt.Columns.Add("Name", typeof(string)).ReadOnly = true;
                dt.Columns.Add("Nickname", typeof(string));

                var nk = this.Nicknames;
                foreach (var Com in ComPorts.Where(p => !string.IsNullOrEmpty(p.DeviceID)))
                    dt.Rows.Add(Com.DeviceID, Com.Port, Com.Name, nk.ContainsKey(Com.DeviceID) ? nk[Com.DeviceID] : "");

                grdList.DataSource = dt;
                grdList.Columns["DeviceID"].Visible = false;
                grdList.Columns["Port"].DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
                grdList.Columns["NickName"].AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
                grdList.AutoResizeColumns();
            };
            this.Activated += (sender, e) =>
            {
                if(grdList.CurrentRow != null)
                    grdList.CurrentCell = grdList.CurrentRow.Cells["Nickname"];
            };
            this.btnReset.Click += (sender, e) =>
            {
                Properties.Settings.Default.Reset();
                this.Close();
            };
            this.btnAccept.Click += (sender, e) =>
            {
                var nk = this.Nicknames;
                foreach (System.Data.DataRow row in (grdList.DataSource as System.Data.DataTable).Rows)
                {
                    var DeviceID = row["DeviceID"].ToString();
                    var Nickname = row["Nickname"].ToString();
                    if (!string.IsNullOrEmpty(Nickname))
                        nk[DeviceID] = Nickname;
                    else if (nk.ContainsKey(DeviceID))
                        nk.Remove(DeviceID);
                }
                Properties.Settings.Default.Nickname = nk.SerializedXML();
                Properties.Settings.Default.Save();
                this.Close();
            };
        }
    }
}
