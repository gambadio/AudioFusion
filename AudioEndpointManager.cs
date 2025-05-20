// AudioEndpointManager.cs

using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text; // For StringBuilder
using System.Threading.Tasks; // Though not strictly used for async methods in this version of the class itself
using NAudio.CoreAudioApi; // For MMDeviceEnumerator, MMDevice, Role, DataFlow, DeviceState, PropertyKeys
using System.Windows; // For MessageBox (used in one place, consider removing or refactoring)

namespace AudioFusion
{
    public static class AudioEndpointManager
    {
        // COM Interfaces for IPolicyConfig (used to set default audio device)
        [ComImport, Guid("F8679F50-850A-41CF-9C72-430F290290C8")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPolicyConfig
        {
            [PreserveSig]
            int GetMixFormat([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, out IntPtr ppFormat);
            [PreserveSig]
            int GetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, bool bDefault, out IntPtr ppFormat);
            [PreserveSig]
            int ResetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName);
            [PreserveSig]
            int SetDeviceFormat([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, IntPtr pEndpointFormat, IntPtr pMixFormat);
            [PreserveSig]
            int GetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, bool bDefault, out long pmftDefaultPeriod, out long pmftMinimumPeriod);
            [PreserveSig]
            int SetProcessingPeriod([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, ref long pmftPeriod);
            [PreserveSig]
            int GetShareMode([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, out IntPtr pShareMode);
            [PreserveSig]
            int SetShareMode([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, IntPtr shareMode);
            [PreserveSig]
            int GetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, ref PropertyKey key, out IntPtr pv);
            [PreserveSig]
            int SetPropertyValue([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, ref PropertyKey key, IntPtr pv);
            [PreserveSig]
            int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, Role role);
            [PreserveSig]
            int SetEndpointVisibility([MarshalAs(UnmanagedType.LPWStr)] string pszDeviceName, bool bVisible);
        }

        [ComImport, Guid("870af99c-171d-4f9e-af0d-e63df40c2bc9")] // CLSID_PolicyConfigClient
        private class PolicyConfigClient { }

        // Minimal PropVariant structure for PKEY_Device_InstanceId (VT_LPWSTR)
        [StructLayout(LayoutKind.Explicit)]
        private struct PropVariant
        {
            [FieldOffset(0)] private ushort valueType;
            [FieldOffset(8)] private IntPtr pointerValue; // VT_LPWSTR

            public string GetString()
            {
                if (valueType == (ushort)VarEnum.VT_LPWSTR && pointerValue != IntPtr.Zero)
                {
                    return Marshal.PtrToStringUni(pointerValue);
                }
                return null;
            }

            // Optional: Method to clear the PropVariant to prevent memory leaks with COM allocations
            [DllImport("Ole32.dll", PreserveSig = false)] // Presig = false means it throws exceptions for HRESULTs
            private static extern void PropVariantClear(ref PropVariant pvar);

            public void Clear()
            {
                if (valueType != (ushort)VarEnum.VT_EMPTY) // Only clear if not already empty
                {
                    PropVariantClear(ref this);
                }
            }
        }

        // WinAPI for broadcasting settings change
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessageTimeout(
            IntPtr hWnd,
            int Msg,
            IntPtr wParam,
            IntPtr lParam,
            int fuFlags,
            int uTimeout,
            IntPtr lpdwResult);

        private const int HWND_BROADCAST = 0xffff;
        private const int WM_SETTINGCHANGE = 0x001A;
        private const int SMTO_ABORTIFHUNG = 0x0002;

        // SetupAPI P/Invokes for enabling/disabling devices
        [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr SetupDiGetClassDevs(
            ref Guid ClassGuid,
            [MarshalAs(UnmanagedType.LPTStr)] string Enumerator,
            IntPtr hwndParent,
            int Flags);

        [DllImport("setupapi.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetupDiDestroyDeviceInfoList(IntPtr DeviceInfoSet);

        [DllImport("setupapi.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetupDiEnumDeviceInfo(
            IntPtr DeviceInfoSet,
            int MemberIndex,
            ref SP_DEVINFO_DATA DeviceInfoData);

        [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetupDiGetDeviceInstanceId(
            IntPtr DeviceInfoSet,
            ref SP_DEVINFO_DATA DeviceInfoData,
            StringBuilder DeviceInstanceId,
            int DeviceInstanceIdSize,
            out int RequiredSize);

        [DllImport("setupapi.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetupDiSetClassInstallParams(
            IntPtr DeviceInfoSet,
            ref SP_DEVINFO_DATA DeviceInfoData,
            ref SP_PROPCHANGE_PARAMS ClassInstallParams,
            int ClassInstallParamsSize);

        [DllImport("setupapi.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetupDiCallClassInstaller(
            uint InstallFunction,
            IntPtr DeviceInfoSet,
            ref SP_DEVINFO_DATA DeviceInfoData);

        [StructLayout(LayoutKind.Sequential)]
        private struct SP_DEVINFO_DATA
        {
            public uint cbSize;
            public Guid ClassGuid;
            public uint DevInst;
            public IntPtr Reserved; // On 64-bit, this is ULONG_PTR, effectively IntPtr
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SP_PROPCHANGE_PARAMS
        {
            public SP_CLASSINSTALL_HEADER ClassInstallHeader;
            public uint StateChange; // DICS_ENABLE, DICS_DISABLE, etc.
            public uint Scope;       // DICS_FLAG_GLOBAL or DICS_FLAG_CONFIGSPECIFIC
            public uint HwProfile;   // Hardware profile ID (0 for current)
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SP_CLASSINSTALL_HEADER
        {
            public uint cbSize;
            public uint InstallFunction; // DIF_PROPERTYCHANGE
        }

        // Flags for SetupDiGetClassDevs
        private const int DIGCF_PRESENT = 0x00000002;
        private const int DIGCF_DEVICEINTERFACE = 0x00000010;
        private const int DIGCF_ALLCLASSES = 0x00000004;

        // Install functions / State changes / Scope flags for SetupAPI
        private const uint DIF_PROPERTYCHANGE = 0x00000012;
        private const uint DICS_ENABLE = 0x00000001;
        private const uint DICS_DISABLE = 0x00000002;
        private const uint DICS_FLAG_GLOBAL = 0x00000001; // Apply change to all hardware profiles

        // Audio device class GUID (Media class, includes audio, video, game controllers)
        // {4d36e96c-e325-11ce-bfc1-08002be10318} - MEDIA
        // {c166523c-fe0c-4a94-a586-f1a80cfbbf3e} - AudioEndpoint (for interfaces, not PnP nodes)
        private static readonly Guid GUID_DEVCLASS_MEDIA = new Guid("{4d36e96c-e325-11ce-bfc1-08002be10318}");

        public static bool SetDefaultAudioEndpoint(string audioEndpointDeviceId, Role role)
        {
            IPolicyConfig policyConfig = null;
            try
            {
                policyConfig = new PolicyConfigClient() as IPolicyConfig;
                if (policyConfig == null)
                {
                    System.Diagnostics.Debug.WriteLine("Failed to create PolicyConfigClient COM object.");
                    return false;
                }

                int hr = policyConfig.SetDefaultEndpoint(audioEndpointDeviceId, role);
                if (hr < 0) // HRESULTs are negative on failure
                {
                    System.Diagnostics.Debug.WriteLine($"SetDefaultEndpoint for '{audioEndpointDeviceId}' (Role: {role}) failed with HRESULT: 0x{hr:X8}");
                    Marshal.ThrowExceptionForHR(hr); // Or handle specific HRESULTs
                    return false;
                }

                System.Diagnostics.Debug.WriteLine($"Successfully set '{audioEndpointDeviceId}' as default for Role: {role}. HRESULT: 0x{hr:X8}");
                
                // Notify all applications about the change. This is important.
                SendMessageTimeout(
                    (IntPtr)HWND_BROADCAST,
                    WM_SETTINGCHANGE,
                    IntPtr.Zero,
                    IntPtr.Zero, // Some suggest (IntPtr)SC_MANAGER_ENUMERATE_SERVICE for lParam, or "Policy" string. Zero seems common.
                    SMTO_ABORTIFHUNG,
                    1000, // Timeout 1 second
                    IntPtr.Zero);
                
                System.Threading.Thread.Sleep(100); // Give a moment for system to process

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Exception in SetDefaultAudioEndpoint: {ex.Message}\n{ex.StackTrace}");
                // Consider specific error handling for COMExceptions, e.g., access denied (E_ACCESSDENIED 0x80070005)
                // which might indicate the need for admin rights.
                if (ex is COMException comEx && (uint)comEx.ErrorCode == 0x80070005)
                {
                     MessageBox.Show("Access denied when setting default audio device. Please try running as administrator.", "Access Denied", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                return false;
            }
            finally
            {
                if (policyConfig != null)
                {
                    Marshal.ReleaseComObject(policyConfig);
                }
            }
        }

        public static string GetDeviceInstanceId(string audioEndpointDeviceId) // Parameter is MMDevice.ID
        {
            MMDeviceEnumerator enumerator = null;
            MMDevice device = null;

            try
            {
                enumerator = new MMDeviceEnumerator();
                device = enumerator.GetDevice(audioEndpointDeviceId);

                if (device == null)
                {
                    System.Diagnostics.Debug.WriteLine($"GetDeviceInstanceId: Device not found for endpoint ID {audioEndpointDeviceId}");
                    return null;
                }

                // Get the instance ID directly using NAudio's PropertyStore
                var instanceId = device.Properties[PropertyKeys.PKEY_Device_InstanceId].Value as string;
                
                if (string.IsNullOrEmpty(instanceId))
                {
                    System.Diagnostics.Debug.WriteLine($"GetDeviceInstanceId: No instance ID found for {device.FriendlyName}.");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"GetDeviceInstanceId for MMDevice.ID {audioEndpointDeviceId} (FriendlyName: {device.FriendlyName}) resolved to PnP Instance ID: {instanceId}");
                return instanceId;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in GetDeviceInstanceId for endpoint ID {audioEndpointDeviceId}: {ex.Message}\n{ex.StackTrace}");
                return null;
            }
            finally
            {
                device?.Dispose();
                enumerator?.Dispose();
            }
        }

        public static bool SetDeviceState(string pnpDeviceInstanceId, bool enable)
        {
            IntPtr deviceInfoSet = IntPtr.Zero;
            try
            {
                Guid devClassGuid = GUID_DEVCLASS_MEDIA; // Use the broader Media class GUID
                deviceInfoSet = SetupDiGetClassDevs(ref devClassGuid, null, IntPtr.Zero, DIGCF_PRESENT);

                if (deviceInfoSet == IntPtr.Zero || deviceInfoSet.ToInt64() == -1) // INVALID_HANDLE_VALUE is -1
                {
                    int error = Marshal.GetLastWin32Error();
                    System.Diagnostics.Debug.WriteLine($"SetDeviceState: SetupDiGetClassDevs failed. Win32Error: {error}");
                    return false;
                }

                SP_DEVINFO_DATA deviceInfoData = new SP_DEVINFO_DATA();
                deviceInfoData.cbSize = (uint)Marshal.SizeOf(typeof(SP_DEVINFO_DATA));

                for (int i = 0; SetupDiEnumDeviceInfo(deviceInfoSet, i, ref deviceInfoData); i++)
                {
                    StringBuilder currentDeviceInstanceId = new StringBuilder(256);
                    if (SetupDiGetDeviceInstanceId(deviceInfoSet, ref deviceInfoData, currentDeviceInstanceId, currentDeviceInstanceId.Capacity, out _))
                    {
                        if (currentDeviceInstanceId.ToString().Equals(pnpDeviceInstanceId, StringComparison.OrdinalIgnoreCase))
                        {
                            // Found the device, now change its state
                            SP_PROPCHANGE_PARAMS propChangeParams = new SP_PROPCHANGE_PARAMS();
                            propChangeParams.ClassInstallHeader.cbSize = (uint)Marshal.SizeOf(typeof(SP_CLASSINSTALL_HEADER));
                            propChangeParams.ClassInstallHeader.InstallFunction = DIF_PROPERTYCHANGE;
                            propChangeParams.StateChange = enable ? DICS_ENABLE : DICS_DISABLE;
                            propChangeParams.Scope = DICS_FLAG_GLOBAL; // Apply to all hardware profiles
                            propChangeParams.HwProfile = 0; // Current hardware profile

                            if (!SetupDiSetClassInstallParams(deviceInfoSet, ref deviceInfoData, ref propChangeParams, Marshal.SizeOf(typeof(SP_PROPCHANGE_PARAMS))))
                            {
                                int error = Marshal.GetLastWin32Error();
                                System.Diagnostics.Debug.WriteLine($"SetDeviceState: SetupDiSetClassInstallParams failed for {pnpDeviceInstanceId}. Win32Error: {error}");
                                return false; // Or continue if you want to try others (though should be unique)
                            }

                            if (!SetupDiCallClassInstaller(DIF_PROPERTYCHANGE, deviceInfoSet, ref deviceInfoData))
                            {
                                int error = Marshal.GetLastWin32Error();
                                System.Diagnostics.Debug.WriteLine($"SetDeviceState: SetupDiCallClassInstaller failed for {pnpDeviceInstanceId}. Win32Error: {error}");
                                // Common errors: ERROR_IN_WOW64 (if 32-bit app on 64-bit OS tries certain things without care)
                                // ERROR_ACCESS_DENIED (5) - needs admin
                                // ERROR_DI_DO_DEFAULT (some devices might need a reboot or have other issues)
                                if (error == 5) // ERROR_ACCESS_DENIED
                                {
                                   MessageBox.Show("Access denied when changing device state. Please try running as administrator.", "Access Denied", MessageBoxButton.OK, MessageBoxImage.Warning);
                                }
                                return false;
                            }
                            System.Diagnostics.Debug.WriteLine($"SetDeviceState: Successfully {(enable ? "enabled" : "disabled")} device {pnpDeviceInstanceId}");
                            return true; // Device found and state change attempted
                        }
                    }
                }
                System.Diagnostics.Debug.WriteLine($"SetDeviceState: Device with PnP Instance ID '{pnpDeviceInstanceId}' not found in MEDIA class.");
                return false; // Device not found
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Exception in SetDeviceState for '{pnpDeviceInstanceId}': {ex.Message}\n{ex.StackTrace}");
                return false;
            }
            finally
            {
                if (deviceInfoSet != IntPtr.Zero && deviceInfoSet.ToInt64() != -1)
                {
                    SetupDiDestroyDeviceInfoList(deviceInfoSet);
                }
            }
        }

        // The SetMicrophoneState method from the original code (which mutes/unmutes)
        // is different from SetDeviceState (which enables/disables at PnP level).
        // Keep it if you need simple mute/unmute functionality.
        // For the Teams scenario, SetDeviceState (disable/enable) is likely more impactful.
        public static bool SetMicrophoneMuteState(string audioEndpointDeviceId, bool mute)
        {
            MMDeviceEnumerator enumerator = null;
            MMDevice device = null;
            try
            {
                enumerator = new MMDeviceEnumerator();
                device = enumerator.GetDevice(audioEndpointDeviceId);
                
                if (device == null)
                {
                    System.Diagnostics.Debug.WriteLine($"SetMicrophoneMuteState: Device not found {audioEndpointDeviceId}");
                    return false;
                }

                if (device.State != DeviceState.Active)
                {
                    System.Diagnostics.Debug.WriteLine($"SetMicrophoneMuteState: Device {device.FriendlyName} is not active.");
                    return false; // Can't mute/unmute if not active
                }

                device.AudioEndpointVolume.Mute = mute;
                System.Diagnostics.Debug.WriteLine($"SetMicrophoneMuteState: Device {device.FriendlyName} mute set to {mute}");
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in SetMicrophoneMuteState for {audioEndpointDeviceId}: {ex.Message}");
                return false;
            }
            finally
            {
                device?.Dispose();
                enumerator?.Dispose();
            }
        }
    }
}