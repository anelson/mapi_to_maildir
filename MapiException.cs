using System;

using MAPI33;

namespace MapiToMaildir
{
	/// <summary>
	/// Summary description for MapiException.
	/// </summary>
	public class MapiException : Exception
	{
		Error _hr;

        public const uint MAPI_E_NOT_FOUND = 0x8004010f;
        public const uint MAPI_E_OUT_OF_MEMORY = 0x8007000e;

		public MapiException(Error hr) : base(String.Format("Error {0}: {1}", (int)hr, hr))
		{
			_hr = hr;
		}
		public MapiException(uint errNo) : base(String.Format("Error 0x{0:x}{1}", errNo, errNo == MAPI_E_OUT_OF_MEMORY ? " (Out of memory; property value likely too long)" : String.Empty)) {
			_hr = Error.ErrorsReturned;
		}

        public Error MapiError {
            get {
                return _hr;
            }
        }
	}
}
