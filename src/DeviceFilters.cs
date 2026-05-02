using System;

namespace SoundDeviceSwitcher
{
    internal static class DeviceFilters
    {
        private static readonly string[] VirtualHints =
        {
            "voicemeeter",
            "vb-audio",
            "virtual",
            "sonar",
            "voicemod",
            "vb-cable",
            "cable output",
            "cable input",
            "aux vaio",
            "vaio",
            "nvidia broadcast"
        };

        public static bool LooksVirtual(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            string lower = value.ToLowerInvariant();
            foreach (string hint in VirtualHints)
            {
                if (lower.IndexOf(hint, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
