namespace HID
{
	/// <summary>
	/// Base class for report types. Simply wraps a byte buffer.
	/// </summary>
	public class Report
	{
        /// <summary> Accessor for the raw byte buffer </summary>
        public byte[] Buffer;

        public Report(int bufferSize)
        {
            Buffer = new byte[bufferSize];
        }

        public Report(byte[] data)
        {
            Buffer = data;
        }
	}
}
