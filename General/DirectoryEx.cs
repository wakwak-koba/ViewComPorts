using System.Linq;

namespace General

{
    public class DictionaryEx<T1, T2> : System.Collections.Generic.Dictionary<T1, T2> where T1 : class where T2 : class
    {
        public DictionaryEx(string SerializedXML)
        {
            try
            {
                using (var dt = this.CreateDataTable())
                {
                    dt.ReadXml(new System.IO.StringReader(SerializedXML));
                    foreach (System.Data.DataRow row in dt.Rows)
                        this.Add(row["Key"] as T1, row["Value"] as T2);
                }
            }
            catch (System.Exception)
            {
                this.Clear();
            }
        }

        private System.Data.DataTable CreateDataTable()
        {
            var dt = new System.Data.DataTable("DictionaryEx");
            dt.PrimaryKey = new[] { dt.Columns.Add("Key", typeof(T1)) };
            dt.Columns.Add("Value", typeof(T2));
            return dt;
        }

        public string SerializedXML()
        {
            using (var dt = this.CreateDataTable())
            {
                foreach (var kv in this)
                    dt.Rows.Add(kv.Key, kv.Value);

                var sw = new System.IO.StringWriter();
                dt.WriteXml(sw);
                return sw.ToString();
            }
        }
    }
}
