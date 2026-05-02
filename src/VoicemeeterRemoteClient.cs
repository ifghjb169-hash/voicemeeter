using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace SoundDeviceSwitcher
{
    internal sealed class VoicemeeterRemoteClient : IDisposable
    {
        private const int RunBanana = 2;

        private IntPtr _library;
        private bool _loggedIn;

        private LoginDelegate _login;
        private LogoutDelegate _logout;
        private RunVoicemeeterDelegate _runVoicemeeter;
        private GetVoicemeeterTypeDelegate _getVoicemeeterType;
        private IsParametersDirtyDelegate _isParametersDirty;
        private InputGetDeviceNumberDelegate _inputGetDeviceNumber;
        private OutputGetDeviceNumberDelegate _outputGetDeviceNumber;
        private InputGetDeviceDescWDelegate _inputGetDeviceDescW;
        private OutputGetDeviceDescWDelegate _outputGetDeviceDescW;
        private SetParameterStringWDelegate _setParameterStringW;
        private GetParameterStringWDelegate _getParameterStringW;
        private SetParameterFloatDelegate _setParameterFloat;

        public string DllPath { get; private set; }

        public void Dispose()
        {
            if (_loggedIn && _logout != null)
            {
                try
                {
                    _logout();
                }
                catch
                {
                    // Ignore shutdown failures.
                }
            }

            _loggedIn = false;

            if (_library != IntPtr.Zero)
            {
                NativeMethods.FreeLibrary(_library);
                _library = IntPtr.Zero;
            }
        }

        public string EnsureConnected()
        {
            EnsureLoaded();

            if (_loggedIn)
            {
                return GetVoicemeeterTypeText();
            }

            int loginResult = _login();
            if (loginResult == 1)
            {
                _runVoicemeeter(RunBanana);
                Thread.Sleep(1800);
                loginResult = _login();
            }

            if (loginResult != 0 && loginResult != -2)
            {
                throw new InvalidOperationException("Voicemeeter Remote login failed. Code=" + loginResult);
            }

            _loggedIn = true;
            PumpParameters(5, 30);
            return GetVoicemeeterTypeText();
        }

        public IList<VoicemeeterDevice> GetInputDevices()
        {
            EnsureConnected();
            return ReadDevices(true);
        }

        public IList<VoicemeeterDevice> GetOutputDevices()
        {
            EnsureConnected();
            return ReadDevices(false);
        }

        public string GetCurrentHardwareInput1()
        {
            EnsureConnected();
            PumpParameters(2, 10);
            return GetStringParameter("Strip[0].Device.Name");
        }

        public string GetCurrentA1Output()
        {
            EnsureConnected();
            PumpParameters(2, 10);
            return GetStringParameter("Bus[0].Device.Name");
        }

        public void SetHardwareInput1(VoicemeeterDevice device)
        {
            SetAndVerifyDeviceParameter("Strip[0].Device.", "Strip[0].Device.Name", device);
        }

        public void SetA1Output(VoicemeeterDevice device)
        {
            SetAndVerifyDeviceParameter("Bus[0].Device.", "Bus[0].Device.Name", device);
        }

        public void RestartAudioEngine()
        {
            EnsureConnected();
            int result = _setParameterFloat("Command.Restart", 1.0f);
            ThrowIfVmError(result, "Cannot restart Voicemeeter audio engine.");
        }

        private void SetDeviceParameter(string prefix, VoicemeeterDevice device)
        {
            if (device == null)
            {
                throw new ArgumentNullException("device");
            }

            EnsureConnected();

            if (!IsSettableDriver(device.Driver))
            {
                throw new InvalidOperationException("Voicemeeter cannot set this device driver type: " + device.Driver);
            }

            string parameter = prefix + device.ParameterSuffix;
            int result = _setParameterStringW(parameter, device.Name);
            ThrowIfVmError(result, "Cannot set Voicemeeter parameter " + parameter + ".");
            PumpParameters(20, 50);
        }

        private string SetAndVerifyDeviceParameter(string prefix, string currentParameter, VoicemeeterDevice device)
        {
            SetDeviceParameter(prefix, device);

            string current = string.Empty;
            for (int i = 0; i < 30; i++)
            {
                PumpParameters(1, 40);
                current = GetStringParameter(currentParameter);
                if (string.Equals(current, device.Name, StringComparison.OrdinalIgnoreCase))
                {
                    return current;
                }
            }

            throw new InvalidOperationException(
                "Voicemeeter accepted the command, but the selected device did not change. Current device is: " +
                (string.IsNullOrWhiteSpace(current) ? "(empty)" : current));
        }

        private IList<VoicemeeterDevice> ReadDevices(bool input)
        {
            var devices = new List<VoicemeeterDevice>();
            int count = input ? _inputGetDeviceNumber() : _outputGetDeviceNumber();
            ThrowIfVmError(count < 0 ? count : 0, "Cannot read Voicemeeter device count.");

            for (int i = 0; i < count; i++)
            {
                int type;
                var name = new StringBuilder(512);
                var hardwareId = new StringBuilder(512);

                int result = input
                    ? _inputGetDeviceDescW(i, out type, name, hardwareId)
                    : _outputGetDeviceDescW(i, out type, name, hardwareId);

                if (result != 0)
                {
                    continue;
                }

                string driver = DriverFromType(type);
                if (!IsSettableDriver(driver))
                {
                    continue;
                }

                devices.Add(new VoicemeeterDevice(i, type, driver, name.ToString(), hardwareId.ToString()));
            }

            return devices;
        }

        private string GetStringParameter(string parameter)
        {
            var value = new StringBuilder(512);
            int result = _getParameterStringW(parameter, value);
            return result == 0 ? value.ToString() : string.Empty;
        }

        private void PumpParameters(int attempts, int delayMs)
        {
            if (_isParametersDirty == null)
            {
                return;
            }

            for (int i = 0; i < attempts; i++)
            {
                int result = _isParametersDirty();
                ThrowIfVmError(result, "Voicemeeter parameter synchronization failed.");
                if (delayMs > 0)
                {
                    Thread.Sleep(delayMs);
                }
            }
        }

        private string GetVoicemeeterTypeText()
        {
            if (_getVoicemeeterType == null)
            {
                return "Voicemeeter";
            }

            int type;
            int result = _getVoicemeeterType(out type);
            if (result != 0)
            {
                return "Voicemeeter";
            }

            switch (type)
            {
                case 1:
                    return "Voicemeeter";
                case 2:
                    return "Voicemeeter Banana";
                case 3:
                    return "Voicemeeter Potato";
                default:
                    return "Voicemeeter";
            }
        }

        private void EnsureLoaded()
        {
            if (_library != IntPtr.Zero)
            {
                return;
            }

            DllPath = FindVoicemeeterRemoteDll();
            _library = NativeMethods.LoadLibrary(DllPath);
            if (_library == IntPtr.Zero)
            {
                int error = Marshal.GetLastWin32Error();
                throw new FileNotFoundException("Cannot load VoicemeeterRemote64.dll. Win32Error=" + error, DllPath);
            }

            _login = GetDelegate<LoginDelegate>("VBVMR_Login");
            _logout = GetDelegate<LogoutDelegate>("VBVMR_Logout");
            _runVoicemeeter = GetDelegate<RunVoicemeeterDelegate>("VBVMR_RunVoicemeeter");
            _getVoicemeeterType = GetDelegate<GetVoicemeeterTypeDelegate>("VBVMR_GetVoicemeeterType");
            _isParametersDirty = GetDelegate<IsParametersDirtyDelegate>("VBVMR_IsParametersDirty");
            _inputGetDeviceNumber = GetDelegate<InputGetDeviceNumberDelegate>("VBVMR_Input_GetDeviceNumber");
            _outputGetDeviceNumber = GetDelegate<OutputGetDeviceNumberDelegate>("VBVMR_Output_GetDeviceNumber");
            _inputGetDeviceDescW = GetDelegate<InputGetDeviceDescWDelegate>("VBVMR_Input_GetDeviceDescW");
            _outputGetDeviceDescW = GetDelegate<OutputGetDeviceDescWDelegate>("VBVMR_Output_GetDeviceDescW");
            _setParameterStringW = GetDelegate<SetParameterStringWDelegate>("VBVMR_SetParameterStringW");
            _getParameterStringW = GetDelegate<GetParameterStringWDelegate>("VBVMR_GetParameterStringW");
            _setParameterFloat = GetDelegate<SetParameterFloatDelegate>("VBVMR_SetParameterFloat");
        }

        private T GetDelegate<T>(string name) where T : class
        {
            IntPtr address = NativeMethods.GetProcAddress(_library, name);
            if (address == IntPtr.Zero)
            {
                throw new MissingMethodException("VoicemeeterRemote64.dll export not found: " + name);
            }

            return Marshal.GetDelegateForFunctionPointer(address, typeof(T)) as T;
        }

        private static string FindVoicemeeterRemoteDll()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string[] candidates =
            {
                Path.Combine(baseDir, "VoicemeeterRemote64.dll"),
                @"C:\Program Files (x86)\VB\Voicemeeter\VoicemeeterRemote64.dll",
                @"C:\Program Files\VB\Voicemeeter\VoicemeeterRemote64.dll"
            };

            foreach (string candidate in candidates)
            {
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }

            throw new FileNotFoundException("VoicemeeterRemote64.dll was not found.");
        }

        private static string DriverFromType(int type)
        {
            switch (type)
            {
                case 1:
                    return "MME";
                case 3:
                    return "WDM";
                case 4:
                    return "KS";
                case 5:
                    return "ASIO";
                default:
                    return "TYPE" + type;
            }
        }

        private static bool IsSettableDriver(string driver)
        {
            return driver == "MME" || driver == "WDM" || driver == "KS" || driver == "ASIO";
        }

        private static void ThrowIfVmError(int result, string message)
        {
            if (result < 0)
            {
                throw new InvalidOperationException(message + " Code=" + result);
            }
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int LoginDelegate();

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int LogoutDelegate();

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int RunVoicemeeterDelegate(int voicemeeterType);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int GetVoicemeeterTypeDelegate(out int voicemeeterType);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int IsParametersDirtyDelegate();

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int InputGetDeviceNumberDelegate();

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int OutputGetDeviceNumberDelegate();

        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Unicode)]
        private delegate int InputGetDeviceDescWDelegate(int index, out int type, StringBuilder name, StringBuilder hardwareId);

        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Unicode)]
        private delegate int OutputGetDeviceDescWDelegate(int index, out int type, StringBuilder name, StringBuilder hardwareId);

        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Unicode)]
        private delegate int SetParameterStringWDelegate(
            [MarshalAs(UnmanagedType.LPStr)] string parameter,
            [MarshalAs(UnmanagedType.LPWStr)] string value);

        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Unicode)]
        private delegate int GetParameterStringWDelegate(
            [MarshalAs(UnmanagedType.LPStr)] string parameter,
            [MarshalAs(UnmanagedType.LPWStr)] StringBuilder value);

        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Ansi)]
        private delegate int SetParameterFloatDelegate(
            [MarshalAs(UnmanagedType.LPStr)] string parameter,
            float value);
    }

    internal static partial class NativeMethods
    {
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr LoadLibrary(string fileName);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Ansi)]
        public static extern IntPtr GetProcAddress(IntPtr module, string procName);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool FreeLibrary(IntPtr module);
    }
}
