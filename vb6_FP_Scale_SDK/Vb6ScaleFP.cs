using System;
using System.Runtime.InteropServices;
using SDK_SC_RFID_Devices;
using RFIDLib;

namespace vb6_FP_Scale_SDK
{
    [ComVisible(true), Guid("4F390D54-1271-4498-8A74-C2B976EE3B60"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IComEvents
    {
        [DispId(0x00000001)]
        void RFIDDeviceFPEvent(string eventType, string args);
    }

    [ComVisible(true), Guid("21535409-7153-45C2-A4CB-EA44A61DF135")]
    public interface IComOjbect
    {
        [DispId(0x10000001)]
        string[] FindDevice();
        [DispId(0x10000002)]
        bool ConnectDevice(string SerialNumber);
        [DispId(0x10000003)]
        bool ConnectDeviceScale(string SerialNumber,string scale);
        [DispId(0x10000004)]
        void DisposeDevice();
        [DispId(0x10000005)]
        void StopScan();
        [DispId(0x10000006)]
        string ScanDevice();
        [DispId(0x10000007)]
        void EnableWait();
        [DispId(0x10000008)]
        void DisbleWait();
        [DispId(0x10000009)]
        string ScanWithWeight();
        [DispId(0x10000010)]
        void loadUserTemplate();
        [DispId(0x10000011)]
        string enrollUserFingerprint(string fname, string lname);
        [DispId(0x10000012)]
        string modifyUserFingerprint(string fname, string lname, string fpTemplate);
        [DispId(0x10000013)]
        void EnableAutoWeight();
        [DispId(0x10000014)]
        void DisableAutoWeight();
    }

    [ClassInterface(ClassInterfaceType.None), ComSourceInterfaces(typeof(IComEvents)), ComVisible(true), Guid("AC3A039A-A45C-4748-9E58-3FDBD580B934")]
    public class Vb6ScaleFP : IComOjbect, IObjectSafety
    {
        [ComVisible(false)]
        public delegate void RFIDDeviceFPEventHandker(string eventType, string args);
        private const int INTERFACESAFE_FOR_UNTRUSTED_CALLER = 1;
        private const int INTERFACESAFE_FOR_UNTRUSTED_DATA = 2;
        private const int S_OK = 0;
        private RFID_Device objDevice;
        public event Vb6ScaleFP.RFIDDeviceFPEventHandker RFIDDeviceFPEvent;

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
            OnDeviceEvent("1", a);
        }
        private void msgAdd(string a, string user)
        {
            OnDeviceEvent("2", a + "-" + user);
        }
        private void errorAdd(string a)
        {
            OnDeviceEvent("3", a);
        }

        private void DevStAdd(string a)
        {
            OnDeviceEvent("4", a);
        }

        private void fpDevStAdd(string a)
        {
            OnDeviceEvent("5", a);
        }

        private void weightAdd(string a, string b)
        {
            OnDeviceEvent("6", a + "-" + b);
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
        private void OnDeviceEvent(string eventType, string args)
        {
            if (this.RFIDDeviceFPEvent != null)
            {
                this.RFIDDeviceFPEvent(eventType, args);
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
