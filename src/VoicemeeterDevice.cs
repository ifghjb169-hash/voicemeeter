namespace SoundDeviceSwitcher
{
    internal sealed class VoicemeeterDevice
    {
        public VoicemeeterDevice(int index, int type, string driver, string name, string hardwareId)
        {
            Index = index;
            Type = type;
            Driver = driver;
            Name = name;
            HardwareId = hardwareId;
        }

        public int Index { get; private set; }
        public int Type { get; private set; }
        public string Driver { get; private set; }
        public string Name { get; private set; }
        public string HardwareId { get; private set; }

        public string ParameterSuffix
        {
            get { return Driver.ToLowerInvariant(); }
        }

        public string Key
        {
            get { return Driver + "|" + Name + "|" + HardwareId; }
        }

        public bool LooksVirtual
        {
            get { return DeviceFilters.LooksVirtual(Name + " " + HardwareId); }
        }
    }
}
