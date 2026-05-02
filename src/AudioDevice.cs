namespace SoundDeviceSwitcher
{
    internal enum SoundDeviceFlow
    {
        Output,
        Input
    }

    internal sealed class AudioDevice
    {
        public AudioDevice(string id, string name, string interfaceName, SoundDeviceFlow flow, int state, bool isDefault)
        {
            Id = id;
            Name = name;
            InterfaceName = interfaceName;
            Flow = flow;
            State = state;
            IsDefault = isDefault;
        }

        public string Id { get; private set; }
        public string Name { get; private set; }
        public string InterfaceName { get; private set; }
        public SoundDeviceFlow Flow { get; private set; }
        public int State { get; private set; }
        public bool IsDefault { get; private set; }

        public string StateText
        {
            get
            {
                switch (State)
                {
                    case 1:
                        return "Ready";
                    case 2:
                        return "Disabled";
                    case 4:
                        return "Not present";
                    case 8:
                        return "Unplugged";
                    default:
                        return "Unknown";
                }
            }
        }

        public string FullName
        {
            get
            {
                if (string.IsNullOrWhiteSpace(InterfaceName) || InterfaceName == Name)
                {
                    return Name;
                }

                return Name + " - " + InterfaceName;
            }
        }
    }
}
