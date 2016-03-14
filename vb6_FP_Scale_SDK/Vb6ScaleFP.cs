using System;
using System.Runtime.InteropServices;
using SDK_SC_RFID_Devices;
using RFIDLib;

namespace vb6_FP_Scale_SDK
{

    [ComVisible(true), Guid("2DBD1BA4-6402-4850-B746-72AAD993946C"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IComEvents
    {
        [DispId(1)]
        void RFIDDeviceFPEvent(int eventType, string args);
    }

    [ComVisible(true), Guid("21535409-7153-45C2-A4CB-EA44A61DF135")]
    public interface IComOjbect
    {
        [DispId(1)]
        string[] FindDevice();
        [DispId(2)]
        bool ConnectDevice(string SerialNumber);
        [DispId(3)]
        bool ConnectDeviceScale(string SerialNumber,string scale);
        [DispId(4)]
        void DisposeDevice();
        [DispId(5)]
        void StopScan();
        [DispId(6)]
        string ScanDevice();
        [DispId(7)]
        void EnableWait();
        [DispId(8)]
        void DisbleWait();
        [DispId(9)]
        string ScanWithWeight();
        [DispId(10)]
        void loadUserTemplate();
        [DispId(11)]
        string enrollUserFingerprint(string fname, string lname);
        [DispId(12)]
        string modifyUserFingerprint(string fname, string lname, string fpTemplate);
        [DispId(13)]
        void EnableAutoWeight();
        [DispId(14)]
        void DisableAutoWeight();
    }

    [ClassInterface(ClassInterfaceType.None), ComSourceInterfaces(typeof(IComEvents)), ComVisible(true), Guid("AC3A039A-A45C-4748-9E58-3FDBD580B934")]
    public class Vb6ScaleFP : IComOjbect, IObjectSafety
    {
        [ComVisible(false)]
        public delegate void RFIDDeviceFPEvent(int eventType, string args);
        private const int INTERFACESAFE_FOR_UNTRUSTED_CALLER = 1;
        private const int INTERFACESAFE_FOR_UNTRUSTED_DATA = 2;
        private const int S_OK = 0;
        private RFID_Device objDevice;
        public event Vb6ScaleFP.RFIDDeviceFPEvent DeviceEvent;

        public Vb6ScaleFP()
		{
            RFID.tagsDelegate = tagAdd;
            RFID.messageDelegate = msgAdd;
            RFID.errorDelegate = errorAdd;
            RFID.weightDelegate = weightAdd;
            RFID.deviceStatusDelegate = DevStAdd;
            RFID.fpstatusDelegate = fpDevStAdd;
		}

        private void tagAdd(string a)
        {
            OnDeviceEvent(1, a);
        }
        private void msgAdd(string a, string user)
        {
            OnDeviceEvent(2, a + "-" + user);
        }
        private void errorAdd(string a)
        {
            OnDeviceEvent(3, a);
        }

        private void DevStAdd(string a)
        {
            OnDeviceEvent(4, a);
        }

        private void fpDevStAdd(string a)
        {
            OnDeviceEvent(5, a);
        }

        private void weightAdd(string a, string b)
        {
            OnDeviceEvent(6, a+"-"+b);
        }

        public string[] FindDevice() { return RFID.FindDevice(); }
        public bool ConnectDevice(string SerialNumber) { return RFID.ConnectDevice(SerialNumber); }
        public bool ConnectDeviceScale(string SerialNumber, string scale) 
        {
            if (scale == "SAT")
                return RFID.ConnectDeviceScale(SerialNumber, RFIDLibWeight.ScaleType.SAT);
            else if (scale == "MAT")
                return RFID.ConnectDeviceScale(SerialNumber, RFIDLibWeight.ScaleType.MAT);
            else if (scale == "CTZ")
                return RFID.ConnectDeviceScale(SerialNumber, RFIDLibWeight.ScaleType.CTZ);
            else
                return false;            
        }
        public void DisposeDevice() { RFID.Dispose(); }
        public void StopScan() { RFID.StopScan(); }
        public string ScanDevice() { return RFID.ScanDevice(); }
        public void EnableWait() { RFID.EnableWait(); }
        public void DisbleWait() { RFID.DisbleWait(); }
        public string ScanWithWeight() { return RFID.ScanWithWeight(); }
        public void loadUserTemplate() { RFID.loadUserTemplate(); }
        public string enrollUserFingerprint(string fname, string lname) { return RFID.enrollUserFingerprint(fname,lname); }
        public string modifyUserFingerprint(string fname, string lname, string fpTemplate) { return RFID.modifyUserFingerprint(fname,lname,fpTemplate); }
        public void EnableAutoWeight() { RFID.EnableAutoWeight(); }
        public void DisableAutoWeight() { RFID.DisableAutoWeight(); }

        [ComVisible(false)]
        private void OnDeviceEvent(int eventType, string args)
        {
            if (this.DeviceEvent != null)
            {
                this.DeviceEvent(eventType, args);
            }
        }
 
        public int GetInterfaceSafetyOptions(ref Guid riid, out int pdwSupportedOptions, out int pdwEnabledOptions)
        {
            pdwSupportedOptions = 3;
            pdwEnabledOptions = 3;
            return 0;
        }
        public int SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions)
        {
            return 0;
        }
    }

    [Guid("F3B4E7FD-0A8E-4AFB-9BD6-1571A556AD1B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    internal interface IObjectSafety
    {
        [PreserveSig]
        int GetInterfaceSafetyOptions(ref Guid riid, out int pdwSupportedOptions, out int pdwEnabledOptions);
        [PreserveSig]
        int SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions);
    }

}
