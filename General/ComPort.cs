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

    public partial class ComPorts : System.Collections.Generic.List<ComPort>
    {
        public ComPorts()
        {
            var r = new System.Text.RegularExpressions.Regex(@"\(COM(?<Port>\d+)\)");

            // WMI
            //using (var pnpEntry = new System.Management.ManagementClass(@"Win32_PnPEntity"))
            //{
            //    this.AddRange(
            //        pnpEntry.GetInstances().OfType<System.Management.ManagementBaseObject>()
            //        .Select(m => new { Name = m.GetPropertyValue("Name")?.ToString(), Manufacturer = m.GetPropertyValue("Manufacturer")?.ToString(), DeviceID = m.GetPropertyValue("DeviceID")?.ToString() })
            //        .Where(m => m.Name != null && r.IsMatch(m.Name))
            //        .Select(m => new ComPort(m.DeviceID, (m.Manufacturer != null && !m.Name.StartsWith(m.Manufacturer) ? m.Manufacturer + " " : "") + m.Name, int.Parse(r.Match(m.Name).Groups["Port"].Value)))
            //        .OrderBy(m => m.Port)
            //    );
            //}

            this.AddRange(
                getDevices("Ports", new[] { SPDRP.HARDWAREID, SPDRP.MFG, SPDRP.FRIENDLYNAME })
                .Select(m => new { InstanceID = m.DeviceInstanceId, DeviceID = m.properties[SPDRP.HARDWAREID],  Manufacturer = m.properties[SPDRP.MFG].FirstOrDefault(), Name = m.properties[SPDRP.FRIENDLYNAME]?.FirstOrDefault()})
                .Where(m => m.Name != null && r.IsMatch(m.Name))
                .Select(m => new ComPort(m.InstanceID, (m.Manufacturer != null && !m.Name.StartsWith(m.Manufacturer) ? m.Manufacturer + " " : "") + m.Name, int.Parse(r.Match(m.Name).Groups["Port"].Value)))
                .OrderBy(m => m.Port)
            );

            var buf = new char[65536];
            var ql = (int)QueryDosDevice(null, buf, (uint)buf.Length);
            r = new System.Text.RegularExpressions.Regex(@"COM[0-9]+");
            this.AddRange(new string(buf, 0, ql).Split('\0')
                .Where(s => r.IsMatch(s))
                .Select(s => new { Name = s, Port = int.Parse(r.Match(s).Value.Substring(3)) })
                .Where(m => this.Where(p => p.Port.Equals(m.Port)).Count() == 0)
                .Select(m => new ComPort(null, m.Name, m.Port))
                .OrderBy(m => m.Port)
            );
        }

        private static System.Collections.Generic.IEnumerable<deviceInfo> getDevices(string className, SPDRP[] properties)
            => getDevices(getGUIDs(className), properties);

        private struct deviceInfo
        {
            public string DeviceInstanceId;
            public System.Collections.Generic.Dictionary<SPDRP, string[]> properties;
        }

        private static System.Collections.Generic.IEnumerable<deviceInfo> getDevices(System.Collections.Generic.IEnumerable<System.Guid> guids, SPDRP[] properties)
        {
            return guids
                .Select(g => SetupDiGetClassDevs(g, null, System.IntPtr.Zero, DIGCF_PRESENT | DIGCF_PROFILE))
                .Where(h => h.ToInt32() != INVALID_HANDLE_VALUE)
                .SelectMany(h => getDevices(h, properties));
        }

        private static System.Collections.Generic.IEnumerable<System.Guid> getGUIDs(string className)
        {
            uint requiredSize;
            var guidArray = new System.Guid[256];
            if (SetupDiClassGuidsFromName(className, ref guidArray[0], (uint)guidArray.Count(), out requiredSize))
                while (requiredSize > 0)
                    yield return guidArray[--requiredSize];
        }

        private static System.Collections.Generic.IEnumerable<deviceInfo> getDevices(System.IntPtr handle, params SPDRP[] properties)
        {
            var da = new SP_DEVINFO_DATA();
            da.cbSize = (uint)System.Runtime.InteropServices.Marshal.SizeOf(da);
            uint i = 0;

            while (SetupDiEnumDeviceInfo(handle, i++, ref da) != 0)
            {
                var buf = new char[1024];
                uint requiredSize;
                if (SetupDiGetDeviceInstanceId(handle, ref da, buf, (uint)buf.Count(), out requiredSize) != 0)
                    yield return new deviceInfo() { DeviceInstanceId = new string(buf, 0, (int)requiredSize - 1), properties = properties.ToDictionary(p => p, p => getProperty(handle, ref da, p)) };
            }
        }

        private static string[] getProperty(System.IntPtr handle, ref SP_DEVINFO_DATA da, SPDRP property)
        {
            var buf = new char[1024];
            uint requiredSize;

            if (SetupDiGetDeviceRegistryProperty(handle, ref da, (uint)property, out _, buf, (uint)buf.Length, out requiredSize))
                return new string(buf, 0, (int)requiredSize / 2 - 1).Split('\0');
            return null;
        }
    }

    public partial class ComPorts
    {
        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true, CharSet = System.Runtime.InteropServices.CharSet.Unicode, EntryPoint = "QueryDosDeviceW")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)]
        private static extern uint QueryDosDevice(
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string lpDeviceName
            , [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPArray)] char[] lpTargetPath
            , [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)] uint ucchMax
        );

        private const int DIGCF_DEFAULT = 0x1;
        private const int DIGCF_PRESENT = 0x2;
        private const int DIGCF_ALLCLASSES = 0x4;
        private const int DIGCF_PROFILE = 0x8;
        private const int DIGCF_DEVICEINTERFACE = 0x10;
        private const int INVALID_HANDLE_VALUE = -1;

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct SP_DEVINFO_DATA
        {
            public uint cbSize;
            public System.Guid ClassGuid;
            public uint DevInst;
            public System.IntPtr Reserved;
        }

        private enum SPDRP
        {
            DEVICEDESC = 0x00000000,
            HARDWAREID = 0x00000001,
            COMPATIBLEIDS = 0x00000002,
            NTDEVICEPATHS = 0x00000003,
            SERVICE = 0x00000004,
            CONFIGURATION = 0x00000005,
            CONFIGURATIONVECTOR = 0x00000006,
            CLASS = 0x00000007,
            CLASSGUID = 0x00000008,
            DRIVER = 0x00000009,
            CONFIGFLAGS = 0x0000000A,
            MFG = 0x0000000B,
            FRIENDLYNAME = 0x0000000C,
            LOCATION_INFORMATION = 0x0000000D,
            PHYSICAL_DEVICE_OBJECT_NAME = 0x0000000E,
            CAPABILITIES = 0x0000000F,
            UI_NUMBER = 0x00000010,
            UPPERFILTERS = 0x00000011,
            LOWERFILTERS = 0x00000012,
            MAXIMUM_PROPERTY = 0x00000013,
        }

        [System.Runtime.InteropServices.DllImport("setupapi.dll", SetLastError = true)]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool SetupDiClassGuidsFromName(
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPStr)] string ClassName
            , ref System.Guid ClassGuidArray1stItem
            , [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)] uint ClassGuidArraySize
            , [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)] out uint RequiredSize);

        [System.Runtime.InteropServices.DllImport("setupapi.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)]
        private static extern uint SetupDiEnumDeviceInfo(System.IntPtr DeviceInfoSet,
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)] uint MemberIndex
            , ref SP_DEVINFO_DATA DeviceInterfaceData);

        [System.Runtime.InteropServices.DllImport("setupapi.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)]
        private static extern uint SetupDiGetDeviceInstanceId(
            System.IntPtr DeviceInfoSet
            , ref SP_DEVINFO_DATA DeviceInfoData
            , [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPArray)] char[] DeviceInstanceId
            , [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)] uint DeviceInstanceIdSize
            , [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)] out uint RequiredSize
        );

        [System.Runtime.InteropServices.DllImport("setupapi.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern System.IntPtr SetupDiGetClassDevs(           // 1st form using a ClassGUID only, with null Enumerator
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPStruct), System.Runtime.InteropServices.In] System.Guid ClassGuid
            , [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr), System.Runtime.InteropServices.In] string Enumerator
            , System.IntPtr hwndParent
            , [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)] uint Flags
        );

        [System.Runtime.InteropServices.DllImport("setupapi.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, SetLastError = true)]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool SetupDiGetDeviceRegistryProperty(
            System.IntPtr DeviceInfoSet
            , ref SP_DEVINFO_DATA DeviceInfoData
            , [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)] uint Property
            , [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)] out uint PropertyRegDataType
            , [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPArray)] char[] PropertyBuffer
            , [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)] uint PropertyBufferSize
            , [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)] out uint RequiredSize
        );
    }
}
