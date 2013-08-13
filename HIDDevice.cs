using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32.SafeHandles;

namespace HID
{
	/// <summary> Generic HID Device - can be extended to customize functionality </summary>
    public class HIDDevice : Win32Usb, IDisposable
    {
        public int ProductID { get; private set; }
        public int VendorID { get; private set; }
        public Guid DeviceClass { get; private set; }
	    public int InputReportLength { get; private set; }
	    public int OutputReportLength { get; private set; }

        public EventHandler<HIDDeviceEventArgs> OnDataReceived = (sender, args) => { };
        public EventHandler<HIDDeviceEventArgs> OnDataSend = (sender, args) => { };
        public EventHandler<HIDDeviceEventArgs> OnDeviceArrived = (sender, args) => { };
        public EventHandler<HIDDeviceEventArgs> OnDeviceRemoved = (sender, args) => { };

        private IntPtr _handle;
        private FileStream _mOFile;
	    private static readonly List<HIDDevice> Devices = new List<HIDDevice>();

        public HIDDevice(int productID, int vendorID)
        {
            DeviceClass = HIDGuid;
            
            ProductID = productID;
            VendorID = vendorID;

            Devices.Add(this);
        }

        public void Dispose()
        {
            if (_mOFile != null)
            {
                _mOFile.Close();
                _mOFile = null;
            }

            // Dispose and finalize, get rid of unmanaged resources
            if (_handle != IntPtr.Zero)
            {
                CloseHandle(_handle);
            }

            GC.SuppressFinalize(this);

            Devices.Remove(this);
        }

        public void SendData(byte[] data)
        {
            // create output report
            var oRep = GenerateOutputReport(data);
            SendData(oRep);
        }

        public void SendData(Report report)
        {
            Write(report);
            OnDataSend.Invoke(this, new HIDDeviceEventArgs { Report = report });
        }

        /// <summary>
        /// Checks the devices that are present at the moment and checks if one of those
        /// is the device you defined by filling in the product id and vendor id.
        /// </summary>
        public void CheckDevicePresent()
        {
            //Mind if the specified device existed before.
            var history = _handle != IntPtr.Zero;
            string devicePath;
            bool isPresent = IsPresent(VendorID, ProductID, out devicePath);

            if (isPresent && !history)
            {
                // device was added
                Initialize(devicePath);
                OnDeviceArrived.Invoke(this, new HIDDeviceEventArgs());
            }
            else if (!isPresent && history)
            {
                // device was removed
                _handle = IntPtr.Zero;
                OnDeviceRemoved.Invoke(this, new HIDDeviceEventArgs());
            }
        }

		/// <summary>
		/// Kicks off an asynchronous read which completes when data is read or when the device
		/// is disconnected. Uses a callback.
		/// </summary>
        private void BeginAsyncRead()
        {
            if (_handle == IntPtr.Zero) return;

            var arrInputReport = new byte[InputReportLength];
             
            // put the buff we used to receive the stuff as the async state then we can get at it when the read completes
            _mOFile.BeginRead(arrInputReport, 0, InputReportLength, ReadCompleted, arrInputReport);
        }

		/// <summary>
		/// Callback for above. Care with this as it will be called on the background thread from the async read
		/// </summary>
		/// <param name="iResult">Async result parameter</param>
        protected void ReadCompleted(IAsyncResult iResult)
        {
            if (_handle == IntPtr.Zero) return;

            // retrieve the read buffer
            var arrBuff = (byte[])iResult.AsyncState;

            try
            {
                // call end read : this throws any exceptions that happened during the read
                _mOFile.EndRead(iResult);

                // Create the input report for the device
                var oInRep = GenerateInputReport(arrBuff);

                // pass the new input report on to the higher level handler
                OnDataReceived.Invoke(this, new HIDDeviceEventArgs { Report = oInRep });

                // when all that is done, kick off another read for the next report
                BeginAsyncRead();
            }
            catch (IOException e)
            {
                Console.WriteLine(e.Message);

                // device was likely unplugged. re-check the connection and invoke OnDeviceRemoved
                CheckDevicePresent();
            }
        }

		/// <summary>
		/// Write an output report to the device.
		/// </summary>
		/// <param name="oOutRep">Output report to write</param>
        protected void Write(Report oOutRep)
		{
            if (_handle == IntPtr.Zero) return;
     
            _mOFile.Write(oOutRep.Buffer, 0, oOutRep.Buffer.Length);
		}

	    /// <summary>
		/// Helper method to return the device path given a DeviceInterfaceData structure and an InfoSet handle.
		/// Used in 'FindDevice' so check that method out to see how to get an InfoSet handle and a DeviceInterfaceData.
		/// </summary>
		/// <param name="hInfoSet">Handle to the InfoSet</param>
		/// <param name="oInterface">DeviceInterfaceData structure</param>
		/// <returns>The device path or null if there was some problem</returns>
		private string GetDevicePath(IntPtr hInfoSet, ref DeviceInterfaceData oInterface)
		{
			uint nRequiredSize = 0;
			// Get the device interface details
			if (!SetupDiGetDeviceInterfaceDetail(hInfoSet, ref oInterface, IntPtr.Zero, 0, ref nRequiredSize, IntPtr.Zero))
			{
				var oDetail = new DeviceInterfaceDetailData {Size = Marshal.SizeOf(typeof (IntPtr)) == 8 ? 8 : 5};

			    if (SetupDiGetDeviceInterfaceDetail(hInfoSet, ref oInterface, ref oDetail, nRequiredSize, ref nRequiredSize, IntPtr.Zero))
				{
					return oDetail.DevicePath;
				}
			}
			return null;
		}

        /// <summary>
		/// Initialises the device
		/// </summary>
		/// <param name="strPath">Path to the device</param>
		private void Initialize(string strPath)
		{
			// Create the file from the device path
            _handle = CreateFile(strPath, GENERIC_READ | GENERIC_WRITE, 0, IntPtr.Zero, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, IntPtr.Zero);

            if ( _handle == InvalidHandleValue)
            {
                _handle = IntPtr.Zero;
				throw HIDDeviceException.GenerateWithWinError("Failed to create device file");
            }

		    IntPtr lpData;

            if (!HidD_GetPreparsedData(_handle, out lpData))
            {
                throw HIDDeviceException.GenerateWithWinError("GetPreparsedData failed");
            }

            HidCaps oCaps;
            HidP_GetCaps(lpData, out oCaps);	// extract the device capabilities from the internal buffer
            InputReportLength = oCaps.InputReportByteLength;	// get the input...
            OutputReportLength = oCaps.OutputReportByteLength;	// ... and output report lengths
                
            _mOFile = new FileStream(new SafeFileHandle(_handle, false), FileAccess.Read | FileAccess.Write, InputReportLength, true);

            BeginAsyncRead();

			HidD_FreePreparsedData(ref lpData);	
		}

        protected bool IsPresent(int vendorID, int productID, out string devicePath)
        {
            var strSearch = string.Format("vid_{0:x4}&pid_{1:x4}", vendorID, productID); 
            var gHid = HIDGuid;

            
            // this gets a list of all HID devices currently connected to the computer (InfoSet)
            var hInfoSet = SetupDiGetClassDevs(ref gHid, null, IntPtr.Zero, DIGCF_DEVICEINTERFACE | DIGCF_PRESENT);	

            // build up a device interface data block
            var oInterface = new DeviceInterfaceData();	
            oInterface.Size = Marshal.SizeOf(oInterface);

            // Now iterate through the InfoSet memory block assigned within Windows in the call to SetupDiGetClassDevs
            // to get device details for each device connected
            int nIndex = 0;
            
            // this gets the device interface information for a device at index 'nIndex' in the memory block
            while (SetupDiEnumDeviceInterfaces(hInfoSet, 0, ref gHid, (uint)nIndex, ref oInterface))	
            {
                // get the device path (see helper method 'GetDevicePath')
                
                devicePath = GetDevicePath(hInfoSet, ref oInterface);

                // do a string search, if we find the VID/PID string then we found our device!
                if (devicePath.IndexOf(strSearch, StringComparison.Ordinal) >= 0)	
                {
                    return true;
                }

                 
                // if we get here, we didn't find our device. So move on to the next one.
                nIndex++;	
            }

            // Before we go, we have to free up the InfoSet memory reserved by SetupDiGetClassDevs
            SetupDiDestroyDeviceInfoList(hInfoSet);
             

            devicePath = "";

            // oops, didn't find our device
            return false;	
        }

        public override bool Equals(object obj)
        {
            if (!(obj is HIDDevice)) return false;
            
            var hid = (HIDDevice) obj;
            return Equals(hid.ProductID, ProductID) && Equals(hid.VendorID, VendorID);
        }

        protected virtual Report GenerateInputReport(byte[] data)
        {
            return new Report(data);
        }

        protected virtual Report GenerateOutputReport(byte[] data)
        {
            return new Report(data);
        }

        /// <summary>
        /// Process Windows messages and check if any HID devices have been added / removed.
        /// </summary>
        /// <param name="m">a ref to Messages, The messages that are thrown from Windows to the application.</param>
        /// <example> This sample shows how to implement this method in your form.
        /// <code> 
        ///protected override void WndProc(ref Message m)
        ///{
        ///    HIDDevice.WndProc(ref m);
        ///    base.WndProc(ref m);
        ///}
        ///</code>
        ///</example>
        public static void WndProc(ref Message m)
        {
            // only respond to USB related messages
            if (m.Msg != WM_DEVICECHANGE) return;

            switch (m.WParam.ToInt32())
            {
                case DEVICE_NODESCHANGED:
                    foreach (var device in Devices)
                    {
                        device.CheckDevicePresent();
                    }
                    break;
            }
        }
    }
}
