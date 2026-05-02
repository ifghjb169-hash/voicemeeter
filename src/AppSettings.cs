using System;
using System.Collections.Generic;
using System.IO;

namespace SoundDeviceSwitcher
{
    internal sealed class AppSettings
    {
        private readonly string _path;
        private readonly Dictionary<string, string> _values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public AppSettings()
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SoundDeviceSwitcher");
            _path = Path.Combine(dir, "settings.ini");
            Load();
        }

        public string SystemOutputDeviceId
        {
            get { return Get("SystemOutputDeviceId"); }
            set { Set("SystemOutputDeviceId", value); }
        }

        public string SystemInputDeviceId
        {
            get { return Get("SystemInputDeviceId"); }
            set { Set("SystemInputDeviceId", value); }
        }

        public string VoicemeeterInputDeviceKey
        {
            get { return Get("VoicemeeterInputDeviceKey"); }
            set { Set("VoicemeeterInputDeviceKey", value); }
        }

        public string VoicemeeterOutputDeviceKey
        {
            get { return Get("VoicemeeterOutputDeviceKey"); }
            set { Set("VoicemeeterOutputDeviceKey", value); }
        }

        public string LanguageCode
        {
            get { return Localizer.NormalizeLanguage(Get("LanguageCode")); }
            set { Set("LanguageCode", Localizer.NormalizeLanguage(value)); }
        }

        public bool HasUserSystemSelection
        {
            get { return !string.IsNullOrWhiteSpace(SystemOutputDeviceId) || !string.IsNullOrWhiteSpace(SystemInputDeviceId); }
        }

        private string Get(string key)
        {
            string value;
            return _values.TryGetValue(key, out value) ? value : string.Empty;
        }

        private void Set(string key, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                _values.Remove(key);
            }
            else
            {
                _values[key] = value;
            }

            Save();
        }

        private void Load()
        {
            if (!File.Exists(_path))
            {
                return;
            }

            foreach (string line in File.ReadAllLines(_path))
            {
                int index = line.IndexOf('=');
                if (index <= 0)
                {
                    continue;
                }

                string key = line.Substring(0, index).Trim();
                string value = line.Substring(index + 1).Trim();
                if (!string.IsNullOrEmpty(key))
                {
                    _values[key] = value;
                }
            }
        }

        private void Save()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_path));
            using (var writer = new StreamWriter(_path, false))
            {
                foreach (var pair in _values)
                {
                    writer.WriteLine(pair.Key + "=" + pair.Value);
                }
            }
        }
    }
}
