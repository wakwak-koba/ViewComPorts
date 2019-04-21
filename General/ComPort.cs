using System.Linq;

namespace General
{
    public class ComPort
    {
        public ComPort(string DeviceID, string Name, int Port)
        {
            this.DeviceID = DeviceID;
            this.Name = Name;
            this.Port = Port;
        }

        public readonly string DeviceID;
        public readonly string Name;
        public readonly int Port;
    }

    public class ComPorts : System.Collections.Generic.List<ComPort>
    {

        public ComPorts()
        {
            using (var pnpEntry = new System.Management.ManagementClass(@"Win32_PnPEntity"))
            {
                var r = new System.Text.RegularExpressions.Regex(@"\(COM(?<Port>\d+)\)");

                this.AddRange(
                    pnpEntry.GetInstances().OfType<System.Management.ManagementBaseObject>()
                    .Select(m => new { Name = m.GetPropertyValue("Name")?.ToString(), Manufacturer = m.GetPropertyValue("Manufacturer")?.ToString(), DeviceID = m.GetPropertyValue("DeviceID")?.ToString() })
                    .Where(m => m.Name != null && r.IsMatch(m.Name))
                    .Select(m => new ComPort(m.DeviceID, (m.Manufacturer != null && !m.Name.StartsWith(m.Manufacturer) ? m.Manufacturer + " " : "") + m.Name, int.Parse(r.Match(m.Name).Groups["Port"].Value)))
                    .OrderBy(m => m.Port)
                );
            }
        }
    }
}
