using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace SoundDeviceSwitcher
{
    internal sealed class CoreAudioController
    {
        private const int ClsctxAll = 23;

        public IList<AudioDevice> GetDevices(SoundDeviceFlow flow)
        {
            try
            {
                return GetDevicesFromCoreAudio(flow);
            }
            catch
            {
                return RegistryAudioDeviceReader.GetDevices(flow, TryGetDefaultDeviceId(flow));
            }
        }

        private IList<AudioDevice> GetDevicesFromCoreAudio(SoundDeviceFlow flow)
        {
            var result = new List<AudioDevice>();
            EDataFlow dataFlow = flow == SoundDeviceFlow.Output ? EDataFlow.eRender : EDataFlow.eCapture;
            IMMDeviceEnumerator enumerator = null;
            IMMDeviceCollection collection = null;

            try
            {
                enumerator = CreateEnumerator();
                string defaultId = GetDefaultDeviceId(enumerator, dataFlow);

                int hr = enumerator.EnumAudioEndpoints(dataFlow, DeviceState.Active, out collection);
                ThrowIfFailed(hr, "Cannot enumerate audio endpoints.");

                uint count;
                ThrowIfFailed(collection.GetCount(out count), "Cannot count audio endpoints.");

                for (uint i = 0; i < count; i++)
                {
                    IMMDevice device = null;
                    try
                    {
                        ThrowIfFailed(collection.Item(i, out device), "Cannot read audio endpoint.");
                        string id;
                        int state;
                        ThrowIfFailed(device.GetId(out id), "Cannot read audio endpoint id.");
                        ThrowIfFailed(device.GetState(out state), "Cannot read audio endpoint state.");

                        string name = ReadDeviceProperty(device, PropertyKeys.PKEY_Device_FriendlyName);
                        string interfaceName = ReadDeviceProperty(device, PropertyKeys.PKEY_DeviceInterface_FriendlyName);
                        if (string.IsNullOrWhiteSpace(interfaceName))
                        {
                            interfaceName = ReadDeviceProperty(device, PropertyKeys.PKEY_Device_DeviceDesc);
                        }

                        result.Add(new AudioDevice(id, name, interfaceName, flow, state, SameDeviceId(id, defaultId)));
                    }
                    finally
                    {
                        ReleaseComObject(device);
                    }
                }
            }
            finally
            {
                ReleaseComObject(collection);
                ReleaseComObject(enumerator);
            }

            return result;
        }

        public void SetDefaultDevice(AudioDevice device)
        {
            if (device == null)
            {
                throw new ArgumentNullException("device");
            }

            SetDefaultDevice(device.Id);
        }

        public void SetDefaultDevice(string deviceId)
        {
            if (string.IsNullOrWhiteSpace(deviceId))
            {
                throw new ArgumentException("Device id is empty.", "deviceId");
            }

            try
            {
                SetDefaultDeviceModern(deviceId);
            }
            catch (InvalidCastException)
            {
                SetDefaultDeviceVista(deviceId);
            }
            catch (COMException ex)
            {
                if ((uint)ex.ErrorCode != 0x80004002)
                {
                    throw;
                }

                SetDefaultDeviceVista(deviceId);
            }
        }

        public bool TryGetEndpointVolumePercent(AudioDevice device, out int volumePercent)
        {
            volumePercent = 0;
            if (device == null || string.IsNullOrWhiteSpace(device.Id))
            {
                return false;
            }

            try
            {
                volumePercent = GetEndpointVolumePercent(device.Id);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public int GetEndpointVolumePercent(string deviceId)
        {
            if (string.IsNullOrWhiteSpace(deviceId))
            {
                throw new ArgumentException("Device id is empty.", "deviceId");
            }

            IMMDeviceEnumerator enumerator = null;
            IMMDevice device = null;
            object endpointVolumeObject = null;
            try
            {
                IAudioEndpointVolume endpointVolume = ActivateEndpointVolume(deviceId, out enumerator, out device, out endpointVolumeObject);
                float scalar;
                ThrowIfFailed(endpointVolume.GetMasterVolumeLevelScalar(out scalar), "Cannot read endpoint volume.");
                int percent = (int)Math.Round(scalar * 100.0f);
                return ClampPercent(percent);
            }
            finally
            {
                ReleaseComObject(endpointVolumeObject);
                ReleaseComObject(device);
                ReleaseComObject(enumerator);
            }
        }

        public void SetEndpointVolumePercent(AudioDevice device, int volumePercent)
        {
            if (device == null)
            {
                throw new ArgumentNullException("device");
            }

            SetEndpointVolumePercent(device.Id, volumePercent);
        }

        public void SetEndpointVolumePercent(string deviceId, int volumePercent)
        {
            if (string.IsNullOrWhiteSpace(deviceId))
            {
                throw new ArgumentException("Device id is empty.", "deviceId");
            }

            int clampedPercent = ClampPercent(volumePercent);
            IMMDeviceEnumerator enumerator = null;
            IMMDevice device = null;
            object endpointVolumeObject = null;
            try
            {
                IAudioEndpointVolume endpointVolume = ActivateEndpointVolume(deviceId, out enumerator, out device, out endpointVolumeObject);
                Guid eventContext = Guid.Empty;
                ThrowIfFailed(endpointVolume.SetMasterVolumeLevelScalar(clampedPercent / 100.0f, ref eventContext), "Cannot set endpoint volume.");
            }
            finally
            {
                ReleaseComObject(endpointVolumeObject);
                ReleaseComObject(device);
                ReleaseComObject(enumerator);
            }
        }

        public bool TryApplyVoicemeeterDefaults(IList<AudioDevice> outputs, IList<AudioDevice> inputs, out string message)
        {
            AudioDevice output;
            AudioDevice input;
            bool changed = TryApplyVoicemeeterDefaults(outputs, inputs, out output, out input);
            var parts = new List<string>();
            parts.Add(output != null ? "Output -> " + output.Name : "VoiceMeeter Input output device not found");
            parts.Add(input != null ? "Input -> " + input.Name : "VoiceMeeter Output/B1 input device not found");
            message = string.Join("; ", parts.ToArray());
            return changed;
        }

        public bool TryApplyVoicemeeterDefaults(IList<AudioDevice> outputs, IList<AudioDevice> inputs, out AudioDevice output, out AudioDevice input)
        {
            output = FindVoicemeeterSystemDevice(outputs, SoundDeviceFlow.Output);
            input = FindVoicemeeterSystemDevice(inputs, SoundDeviceFlow.Input);
            bool changed = false;

            if (output != null)
            {
                SetDefaultDevice(output);
                changed = true;
            }

            if (input != null)
            {
                SetDefaultDevice(input);
                changed = true;
            }

            return changed;
        }

        private static AudioDevice FindVoicemeeterSystemDevice(IList<AudioDevice> devices, SoundDeviceFlow flow)
        {
            string[] requiredNames = flow == SoundDeviceFlow.Output
                ? new[] { "voicemeeter input", "voicemeter input" }
                : new[] { "voicemeeter output", "voicemeter output", "voicemeeter b1", "voicemeter b1" };

            foreach (AudioDevice device in devices)
            {
                string name = NormalizeDeviceText(device.Name);
                string value = NormalizeDeviceText(device.Name + " " + device.InterfaceName);
                if (ContainsAny(name, requiredNames) && IsMainVoicemeeterVaio(value))
                {
                    return device;
                }
            }

            foreach (AudioDevice device in devices)
            {
                string name = NormalizeDeviceText(device.Name);
                string value = NormalizeDeviceText(device.Name + " " + device.InterfaceName);
                if (ContainsAny(name, requiredNames) && IsMainVoicemeeterEndpoint(value))
                {
                    return device;
                }
            }

            return null;
        }

        private static string NormalizeDeviceText(string value)
        {
            return (value ?? string.Empty).ToLowerInvariant();
        }

        private static bool ContainsAny(string value, string[] candidates)
        {
            foreach (string candidate in candidates)
            {
                if (value.Contains(candidate))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsMainVoicemeeterVaio(string value)
        {
            return IsMainVoicemeeterEndpoint(value)
                && (value.Contains("voicemeeter vaio") || value.Contains("voicemeter vaio"));
        }

        private static bool IsMainVoicemeeterEndpoint(string value)
        {
            return (value.Contains("voicemeeter") || value.Contains("voicemeter"))
                && !value.Contains("aux");
        }

        private static void SetDefaultForRole(IPolicyConfig policyConfig, string deviceId, ERole role)
        {
            int hr = policyConfig.SetDefaultEndpoint(deviceId, role);
            ThrowIfFailed(hr, "Cannot set default audio endpoint.");
        }

        private static void SetDefaultForRole(IPolicyConfigVista policyConfig, string deviceId, ERole role)
        {
            int hr = policyConfig.SetDefaultEndpoint(deviceId, role);
            ThrowIfFailed(hr, "Cannot set default audio endpoint.");
        }

        private static void SetDefaultDeviceModern(string deviceId)
        {
            object policyConfigObject = null;
            try
            {
                policyConfigObject = CreatePolicyConfigClient(new Guid("870AF99C-171D-4F9E-AF0D-E63DF40C2BC9"));
                IPolicyConfig policyConfig = (IPolicyConfig)policyConfigObject;
                SetDefaultForRole(policyConfig, deviceId, ERole.eConsole);
                SetDefaultForRole(policyConfig, deviceId, ERole.eMultimedia);
                SetDefaultForRole(policyConfig, deviceId, ERole.eCommunications);
            }
            finally
            {
                ReleaseComObject(policyConfigObject);
            }
        }

        private static void SetDefaultDeviceVista(string deviceId)
        {
            object policyConfigObject = null;
            try
            {
                policyConfigObject = CreatePolicyConfigClient(new Guid("294935CE-F637-4E7C-A41B-AB255460B862"));
                IPolicyConfigVista policyConfig = (IPolicyConfigVista)policyConfigObject;
                SetDefaultForRole(policyConfig, deviceId, ERole.eConsole);
                SetDefaultForRole(policyConfig, deviceId, ERole.eMultimedia);
                SetDefaultForRole(policyConfig, deviceId, ERole.eCommunications);
            }
            finally
            {
                ReleaseComObject(policyConfigObject);
            }
        }

        private static IMMDeviceEnumerator CreateEnumerator()
        {
            Type type = Type.GetTypeFromCLSID(new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E"), true);
            return (IMMDeviceEnumerator)Activator.CreateInstance(type);
        }

        private static object CreatePolicyConfigClient(Guid classId)
        {
            Type type = Type.GetTypeFromCLSID(classId, true);
            return Activator.CreateInstance(type);
        }

        private static IAudioEndpointVolume ActivateEndpointVolume(string deviceId, out IMMDeviceEnumerator enumerator, out IMMDevice device, out object endpointVolumeObject)
        {
            enumerator = null;
            device = null;
            endpointVolumeObject = null;

            enumerator = CreateEnumerator();
            ThrowIfFailed(enumerator.GetDevice(deviceId, out device), "Cannot open audio endpoint.");

            Guid iid = typeof(IAudioEndpointVolume).GUID;
            ThrowIfFailed(device.Activate(ref iid, ClsctxAll, IntPtr.Zero, out endpointVolumeObject), "Cannot activate endpoint volume.");
            return (IAudioEndpointVolume)endpointVolumeObject;
        }

        private static string GetDefaultDeviceId(IMMDeviceEnumerator enumerator, EDataFlow flow)
        {
            IMMDevice defaultDevice = null;
            try
            {
                int hr = enumerator.GetDefaultAudioEndpoint(flow, ERole.eMultimedia, out defaultDevice);
                if (hr != 0 || defaultDevice == null)
                {
                    return string.Empty;
                }

                string id;
                hr = defaultDevice.GetId(out id);
                return hr == 0 ? id : string.Empty;
            }
            finally
            {
                ReleaseComObject(defaultDevice);
            }
        }

        private static string TryGetDefaultDeviceId(SoundDeviceFlow flow)
        {
            IMMDeviceEnumerator enumerator = null;
            try
            {
                enumerator = CreateEnumerator();
                EDataFlow dataFlow = flow == SoundDeviceFlow.Output ? EDataFlow.eRender : EDataFlow.eCapture;
                return GetDefaultDeviceId(enumerator, dataFlow);
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ReleaseComObject(enumerator);
            }
        }

        private static string ReadDeviceProperty(IMMDevice device, PROPERTYKEY key)
        {
            IPropertyStore store = null;
            PropVariant value = new PropVariant();
            try
            {
                int hr = device.OpenPropertyStore(0, out store);
                if (hr != 0 || store == null)
                {
                    return string.Empty;
                }

                hr = store.GetValue(ref key, out value);
                if (hr != 0)
                {
                    return string.Empty;
                }

                return value.GetString();
            }
            finally
            {
                NativeMethods.PropVariantClear(ref value);
                ReleaseComObject(store);
            }
        }

        private static bool SameDeviceId(string left, string right)
        {
            return string.Equals(left ?? string.Empty, right ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }

        private static int ClampPercent(int value)
        {
            if (value < 0)
            {
                return 0;
            }

            if (value > 100)
            {
                return 100;
            }

            return value;
        }

        private static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }

        private static void ThrowIfFailed(int hr, string message)
        {
            if (hr < 0)
            {
                throw new InvalidOperationException(message + " HRESULT=0x" + hr.ToString("X8"), Marshal.GetExceptionForHR(hr));
            }
        }
    }

    internal static class RegistryAudioDeviceReader
    {
        private const string BaseKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio";
        private static readonly Regex EndpointIdRegex = new Regex(@"\{0\.0\.[01]\.00000000\}\.\{[0-9a-fA-F\-]{36}\}", RegexOptions.Compiled);

        public static IList<AudioDevice> GetDevices(SoundDeviceFlow flow, string defaultId)
        {
            var devices = new List<AudioDevice>();
            string flowKey = flow == SoundDeviceFlow.Output ? "Render" : "Capture";

            using (RegistryKey root = Registry.LocalMachine.OpenSubKey(BaseKey + "\\" + flowKey))
            {
                if (root == null)
                {
                    return devices;
                }

                foreach (string endpointKeyName in root.GetSubKeyNames())
                {
                    using (RegistryKey endpointKey = root.OpenSubKey(endpointKeyName))
                    using (RegistryKey propertiesKey = endpointKey == null ? null : endpointKey.OpenSubKey("Properties"))
                    {
                        if (endpointKey == null || propertiesKey == null)
                        {
                            continue;
                        }

                        int state = ReadInt(endpointKey.GetValue("DeviceState"), 0);
                        if (state != DeviceState.Active)
                        {
                            continue;
                        }

                        string name = ReadString(propertiesKey, "{a45c254e-df1c-4efd-8020-67d146a850e0},2");
                        if (string.IsNullOrWhiteSpace(name))
                        {
                            name = ReadString(propertiesKey, "{a45c254e-df1c-4efd-8020-67d146a850e0},14");
                        }

                        string interfaceName = ReadString(propertiesKey, "{b3f8fa53-0004-438e-9003-51a46e139bfc},6");
                        if (string.IsNullOrWhiteSpace(interfaceName))
                        {
                            interfaceName = ReadString(propertiesKey, "{9dad2fed-3c19-4cde-b3c9-1bd56be25698},0");
                        }

                        if (string.IsNullOrWhiteSpace(name))
                        {
                            name = endpointKeyName;
                        }

                        string id = ReadEndpointId(propertiesKey, flow, endpointKeyName);
                        devices.Add(new AudioDevice(id, name, interfaceName, flow, state, SameDeviceId(id, defaultId)));
                    }
                }
            }

            return devices;
        }

        private static string ReadEndpointId(RegistryKey propertiesKey, SoundDeviceFlow flow, string endpointKeyName)
        {
            string deviceInstance = ReadString(propertiesKey, "{9c119480-ddc2-4954-a150-5bd240d454ad},2");
            if (!string.IsNullOrWhiteSpace(deviceInstance))
            {
                Match match = EndpointIdRegex.Match(deviceInstance);
                if (match.Success)
                {
                    return match.Value;
                }
            }

            string prefix = flow == SoundDeviceFlow.Output ? "{0.0.0.00000000}." : "{0.0.1.00000000}.";
            return prefix + endpointKeyName;
        }

        private static string ReadString(RegistryKey key, string name)
        {
            object value = key.GetValue(name);
            if (value == null)
            {
                return string.Empty;
            }

            string text = value as string;
            if (text != null)
            {
                return text;
            }

            string[] strings = value as string[];
            if (strings != null && strings.Length > 0)
            {
                return strings[0];
            }

            return value.ToString();
        }

        private static int ReadInt(object value, int fallback)
        {
            if (value is int)
            {
                return (int)value;
            }

            if (value == null)
            {
                return fallback;
            }

            int parsed;
            return int.TryParse(value.ToString(), out parsed) ? parsed : fallback;
        }

        private static bool SameDeviceId(string left, string right)
        {
            return string.Equals(left ?? string.Empty, right ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }
    }

    internal enum EDataFlow
    {
        eRender = 0,
        eCapture = 1,
        eAll = 2
    }

    internal enum ERole
    {
        eConsole = 0,
        eMultimedia = 1,
        eCommunications = 2
    }

    internal static class DeviceState
    {
        public const int Active = 0x00000001;
    }

    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    internal class MMDeviceEnumerator
    {
    }

    [ComImport]
    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceEnumerator
    {
        [PreserveSig]
        int EnumAudioEndpoints(EDataFlow dataFlow, int dwStateMask, out IMMDeviceCollection ppDevices);

        [PreserveSig]
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppEndpoint);

        [PreserveSig]
        int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string pwstrId, out IMMDevice ppDevice);

        [PreserveSig]
        int RegisterEndpointNotificationCallback(IntPtr pClient);

        [PreserveSig]
        int UnregisterEndpointNotificationCallback(IntPtr pClient);
    }

    [ComImport]
    [Guid("0BD7A1BE-7A1A-44DB-8397-C0F1AF5EE574")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceCollection
    {
        [PreserveSig]
        int GetCount(out uint pcDevices);

        [PreserveSig]
        int Item(uint nDevice, out IMMDevice ppDevice);
    }

    [ComImport]
    [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDevice
    {
        [PreserveSig]
        int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);

        [PreserveSig]
        int OpenPropertyStore(int stgmAccess, out IPropertyStore ppProperties);

        [PreserveSig]
        int GetId([MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);

        [PreserveSig]
        int GetState(out int pdwState);
    }

    [ComImport]
    [Guid("5CDF2C82-841E-4546-9722-0CF74078229A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioEndpointVolume
    {
        [PreserveSig]
        int RegisterControlChangeNotify(IntPtr pNotify);

        [PreserveSig]
        int UnregisterControlChangeNotify(IntPtr pNotify);

        [PreserveSig]
        int GetChannelCount(out uint pnChannelCount);

        [PreserveSig]
        int SetMasterVolumeLevel(float fLevelDB, ref Guid pguidEventContext);

        [PreserveSig]
        int SetMasterVolumeLevelScalar(float fLevel, ref Guid pguidEventContext);

        [PreserveSig]
        int GetMasterVolumeLevel(out float pfLevelDB);

        [PreserveSig]
        int GetMasterVolumeLevelScalar(out float pfLevel);

        [PreserveSig]
        int SetChannelVolumeLevel(uint nChannel, float fLevelDB, ref Guid pguidEventContext);

        [PreserveSig]
        int SetChannelVolumeLevelScalar(uint nChannel, float fLevel, ref Guid pguidEventContext);

        [PreserveSig]
        int GetChannelVolumeLevel(uint nChannel, out float pfLevelDB);

        [PreserveSig]
        int GetChannelVolumeLevelScalar(uint nChannel, out float pfLevel);

        [PreserveSig]
        int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, ref Guid pguidEventContext);

        [PreserveSig]
        int GetMute([MarshalAs(UnmanagedType.Bool)] out bool pbMute);

        [PreserveSig]
        int GetVolumeStepInfo(out uint pnStep, out uint pnStepCount);

        [PreserveSig]
        int VolumeStepUp(ref Guid pguidEventContext);

        [PreserveSig]
        int VolumeStepDown(ref Guid pguidEventContext);

        [PreserveSig]
        int QueryHardwareSupport(out uint pdwHardwareSupportMask);

        [PreserveSig]
        int GetVolumeRange(out float pflVolumeMindB, out float pflVolumeMaxdB, out float pflVolumeIncrementdB);
    }

    [ComImport]
    [Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPropertyStore
    {
        [PreserveSig]
        int GetCount(out uint cProps);

        [PreserveSig]
        int GetAt(uint iProp, out PROPERTYKEY pkey);

        [PreserveSig]
        int GetValue(ref PROPERTYKEY key, out PropVariant pv);

        [PreserveSig]
        int SetValue(ref PROPERTYKEY key, ref PropVariant propvar);

        [PreserveSig]
        int Commit();
    }

    [ComImport]
    [Guid("870AF99C-171D-4F9E-AF0D-E63DF40C2BC9")]
    internal class PolicyConfigClient
    {
    }

    [ComImport]
    [Guid("F8679F50-850A-41CF-9C72-430F290290C8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPolicyConfig
    {
        [PreserveSig]
        int GetMixFormat([MarshalAs(UnmanagedType.LPWStr)] string deviceId, IntPtr ppFormat);

        [PreserveSig]
        int GetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string deviceId, int defaultFormat, IntPtr ppFormat);

        [PreserveSig]
        int ResetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string deviceId);

        [PreserveSig]
        int SetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string deviceId, IntPtr endpointFormat, IntPtr mixFormat);

        [PreserveSig]
        int GetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string deviceId, int defaultPeriod, IntPtr defaultPeriodValue, IntPtr minimumPeriodValue);

        [PreserveSig]
        int SetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string deviceId, IntPtr period);

        [PreserveSig]
        int GetShareMode([MarshalAs(UnmanagedType.LPWStr)] string deviceId, IntPtr mode);

        [PreserveSig]
        int SetShareMode([MarshalAs(UnmanagedType.LPWStr)] string deviceId, IntPtr mode);

        [PreserveSig]
        int GetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string deviceId, ref PROPERTYKEY key, IntPtr propvar);

        [PreserveSig]
        int SetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string deviceId, ref PROPERTYKEY key, ref PropVariant propvar);

        [PreserveSig]
        int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string deviceId, ERole role);

        [PreserveSig]
        int SetEndpointVisibility([MarshalAs(UnmanagedType.LPWStr)] string deviceId, int visible);
    }

    [ComImport]
    [Guid("568B9108-44BF-40B4-9006-86AFE5B5A620")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPolicyConfigVista
    {
        [PreserveSig]
        int GetMixFormat([MarshalAs(UnmanagedType.LPWStr)] string deviceId, IntPtr ppFormat);

        [PreserveSig]
        int GetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string deviceId, int defaultFormat, IntPtr ppFormat);

        [PreserveSig]
        int SetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string deviceId, IntPtr endpointFormat, IntPtr mixFormat);

        [PreserveSig]
        int GetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string deviceId, int defaultPeriod, IntPtr defaultPeriodValue, IntPtr minimumPeriodValue);

        [PreserveSig]
        int SetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string deviceId, IntPtr period);

        [PreserveSig]
        int GetShareMode([MarshalAs(UnmanagedType.LPWStr)] string deviceId, IntPtr mode);

        [PreserveSig]
        int SetShareMode([MarshalAs(UnmanagedType.LPWStr)] string deviceId, IntPtr mode);

        [PreserveSig]
        int GetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string deviceId, ref PROPERTYKEY key, IntPtr propvar);

        [PreserveSig]
        int SetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string deviceId, ref PROPERTYKEY key, ref PropVariant propvar);

        [PreserveSig]
        int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string deviceId, ERole role);

        [PreserveSig]
        int SetEndpointVisibility([MarshalAs(UnmanagedType.LPWStr)] string deviceId, int visible);
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    internal struct PROPERTYKEY
    {
        public Guid fmtid;
        public int pid;

        public PROPERTYKEY(Guid fmtid, int pid)
        {
            this.fmtid = fmtid;
            this.pid = pid;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct PropVariant
    {
        public ushort vt;
        public ushort wReserved1;
        public ushort wReserved2;
        public ushort wReserved3;
        public IntPtr p;
        public int p2;

        public string GetString()
        {
            if (vt == 31 && p != IntPtr.Zero)
            {
                return Marshal.PtrToStringUni(p);
            }

            return string.Empty;
        }
    }

    internal static class PropertyKeys
    {
        public static readonly PROPERTYKEY PKEY_Device_FriendlyName =
            new PROPERTYKEY(new Guid("A45C254E-DF1C-4EFD-8020-67D146A850E0"), 14);

        public static readonly PROPERTYKEY PKEY_Device_DeviceDesc =
            new PROPERTYKEY(new Guid("A45C254E-DF1C-4EFD-8020-67D146A850E0"), 2);

        public static readonly PROPERTYKEY PKEY_DeviceInterface_FriendlyName =
            new PROPERTYKEY(new Guid("026E516E-B814-414B-83CD-856D6FEF4822"), 2);
    }

    internal static partial class NativeMethods
    {
        [DllImport("Ole32.dll")]
        public static extern int PropVariantClear(ref PropVariant pvar);
    }
}
