using System;
using System.Collections;
using System.IO;
using System.Text.RegularExpressions;

using MAPI33;
using MAPI33.MapiTypes;
using MapiGuts = __MAPI33__INTERNALS__;

namespace MapiToMaildir
{
	/// <summary>
	/// Static utility class w/ a few helper functions for MAPI33
	/// </summary>
	internal class MapiUtils {

		/// <summary>
		/// Gets the uncompressed RTF body of a message
		/// </summary>
		/// <param name="msg"></param>
		/// <returns></returns>
		public static byte[] GetBodyRtf(IMessage msg) {
			byte[] bodyRtf = null;

			IUnknown unk;
			Error hr = msg.OpenProperty(Tags.PR_RTF_COMPRESSED, IStream.IID, 0, 0, out unk);
			if (hr == Error.Success) {
				using (unk) {
					IStream compressedStream = (IStream)unk;
					using (compressedStream) {
						IStream uncompStream = null;

						hr = MAPI33.EDK.RTF.WrapCompressedRTFStream(compressedStream, 0, out uncompStream);
						if (hr == Error.Success) {
							using (uncompStream) {
								int bytesRead = 0;
								ArrayList streamContents = new ArrayList();
								byte[] buffer = new byte[1024];

								do {
									uncompStream.Read(buffer, (uint)buffer.Length, out bytesRead);
									if (bytesRead == buffer.Length) {
										streamContents.AddRange(buffer);
									} else {
										for (int idx = 0; idx < (int)bytesRead; idx++) {
											streamContents.Add(buffer[idx]);
										}
									}
								} while (bytesRead > 0);

								bodyRtf = (byte[])streamContents.ToArray(typeof(byte));
							}
						}
					}
				}
			}

			return bodyRtf;
		}

		/// <summary>
		/// Often, a message received as HTML will be exposed by MAPI as RTF, with some horrific RTF tags decorating the
		/// HTML.  Obviously, only Outlook knows how to render this, which isn't cool.  This method tries to detect this, and remove
		/// the HTML wrapper tags.  If it looks like normal RTF and not HTML wrapper RTF, returns null.
		/// 
		/// Assumes the RTF is UTF-8 encoded, which seems a reasonable assumption.
		/// </summary>
		/// <param name="rtf"></param>
		/// <returns></returns>
		public static String BodyRtfToHtml(byte[] rtf) {
			String rtfString = new String(System.Text.Encoding.UTF8.GetChars(rtf));

			if (rtfString.StartsWith(@"{\rtf1")) {
				//This is def an RTF

				//Look for RTF tag \fromhtml indicating this is based on HTML
				if (Regex.IsMatch(rtfString, Regex.Escape(@"\fromhtml"))) {
					//This has the teltale RTF tag.  Try to strip them all out, leaving only the
					//HTML behind.
					//
#if DEBUG
					using (System.IO.StreamWriter sw = System.IO.File.CreateText("g:\\temp\\rtfhtml.rtf")) {
						sw.Write(rtfString);
					}
#endif

					//Remove all newlines.  The real line breaks are encoded as \par tags
					rtfString = rtfString.Replace(Environment.NewLine, "");
					rtfString = Regex.Replace(rtfString, @"(?<!\\)\\par(?>\s)", Environment.NewLine);

					//Remove all the RTF blocks delimited by \htmlrtf and \htmlrtf0.  These blocks contain
					//RTF cruft and can be safely deleted
					//Counter-intuitively, the SingleLine option causes the . class to match newlines as well as other chars,
					//so specify this option to ensure the expression matches multi-line htmlrtf...htmlrtf0 blocks
					const String RTF_BLOCK = @"(?<!\\)\\htmlrtf.*?(?<!\\)\\htmlrtf0";
					while (Regex.IsMatch(rtfString, RTF_BLOCK, RegexOptions.Singleline)) {
						rtfString = Regex.Replace(rtfString, RTF_BLOCK, "", RegexOptions.Singleline);
					}

					//Remove all RTF between \pntext and the nearest block end.  From empirical study and 
					//http://www.wischik.com/lu/programmer/mapi_utils.html, even literal text following \pntext is
					//usually cruft, for example a * character after an LI element
					rtfString = Regex.Replace(rtfString, @"(?<!\\)\\pntext.*?(?=\})", "");

					//Remove all un-escaped { and } characters, leaving the contents of the blocks they delimit
					rtfString = Regex.Replace(rtfString, @"(?<!\\)[\{\}]", "");

					//Remove everything that looks like an RTF tag
					rtfString = Regex.Replace(rtfString, @"(?<!\\)\\(?:\*|'?[a-z\-0-9]+)", "", RegexOptions.Multiline);

					//Un-escape the curly braces escaped by RTF
					rtfString = rtfString.Replace(@"\{", "{");
					rtfString = rtfString.Replace(@"\}", "}");

					//Remove everything before the first HTML tag (usually but not necessarily <HTML>)
					int htmlidx = rtfString.IndexOf("<");
					if (htmlidx == -1) {
						//Can't really be HTMl now can it?
						return null;
					}

					rtfString = rtfString.Substring(htmlidx);

#if DEBUG
					using (System.IO.StreamWriter sw = System.IO.File.CreateText("g:\\temp\\rtfhtml.html")) {
						sw.Write(rtfString);
					}
#endif

					return rtfString;
				}
			}

			return null;
		}

		public static String GetStringProperty(MapiGuts.MAPIProp obj, Tags propId) {
			//Helper to get a single string property from a MAPI object
			Value val = MapiUtils.GetProperty(obj, propId);
			if (val != null) {
				return val.ToString();
			} else {
				return null;
			}
        }

        public static String GetLongStringProperty(MapiGuts.MAPIProp obj, Tags propId) {
            //Gets a string that is potentially long, and therefore must be retrieved
            //using a stream.

			IUnknown unk;
            Error hr = obj.OpenProperty(propId, IStream.IID, 0, 0, out unk);
			if (hr == Error.Success) {
				using (unk) {
					IStream stream = (IStream)unk;
					using (stream) {
                        using (MemoryStream properyContents = new MemoryStream()) {

                            byte[] buffer = new byte[1024];
                            int read = 0;

                            do {
                                hr = stream.Read(buffer, (uint)buffer.Length, out read);
                                properyContents.Write(buffer, 0, read);
                            } while (hr == Error.Success && read > 0);

                            //Seek back to the start of the stream, and read it as text
                            properyContents.Seek(0, SeekOrigin.Begin);

                            //UTF-8 seems reasonable, at least for my email.
                            using (StreamReader rdr = new StreamReader(properyContents, System.Text.Encoding.UTF8)) {
                                return rdr.ReadToEnd();
                            }
                        }
                    }
                }
            }

            return String.Format("Error getting {0}: {1:X}", propId, hr);
        }

		public static byte[] GetBinaryProperty(MapiGuts.MAPIProp obj, Tags propId) {
			//Helper to get a single string property from a MAPI object
			Value val = MapiUtils.GetProperty(obj, propId);
			if (val != null) {
				return ((MapiBinary)val).Value;
			} else {
				return null;
			}
		}

		public static ENTRYID GetEntryIdProperty(MapiGuts.MAPIProp obj, Tags propId) {
			//Helper to get a single Entry ID property from a MAPI object
			Value val = MapiUtils.GetProperty(obj, propId);
			if (val != null) {
				return ENTRYID.Create(val);
			} else {
				return null;
			}
		}

		public static DateTime GetSysTimeProperty(MapiGuts.MAPIProp obj, Tags propId) {
			Value val = MapiUtils.GetProperty(obj, propId);
			if (val != null) {
				return ((MapiSysTime)val).Value;
			} else {
				return DateTime.MinValue;
			}
		}

		public static DateTime GetAppTimeProperty(MapiGuts.MAPIProp obj, Tags propId) {
			Value val = MapiUtils.GetProperty(obj, propId);
			if (val != null) {
				return ((MapiAppTime)val).Value;
			} else {
				return DateTime.MinValue;
			}
		}

		public static int GetLongProperty(MapiGuts.MAPIProp obj, Tags propId) {
			Value val = MapiUtils.GetProperty(obj, propId);
			if (val != null) {
				return ((MapiInt32)val).Value;
			} else {
				return Int32.MinValue;
			}
		}

		public static Value GetProperty(MapiGuts.MAPIProp obj, Tags propId) {
			//Helper to get a single property from a MAPI object
			Tags[] propIds = new Tags[] {propId};
			Value[] values;

			try {
				Error hr = obj.GetProps(propIds, 0, out values);
				if (hr != Error.Success) {
					if (hr == Error.ErrorsReturned) {
						//There was an error getting the property.  The error code is returned instead
						uint error = (uint)((MapiError)values[0]).Value;

						if (error == MapiException.MAPI_E_NOT_FOUND) {
							//Property not found
							return null;
						}
						throw new MapiException(error);
					}

					throw new MapiException(hr);
				}

				return values[0];
			} catch (Exception e) {
				return new MapiString(Tags.PR_DISPLAY_NAME, String.Format("Exception: {0}", e.Message));
			}
		}
	}
}
