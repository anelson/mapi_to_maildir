using System;
using System.IO;
using System.Collections.Specialized;
using System.Collections;
using System.Text.RegularExpressions;

using MAPI33;
using MAPI33.MapiTypes;
using MapiGuts = __MAPI33__INTERNALS__;

namespace MapiToMaildir
{
	/// <summary>
	/// Class to write a maildir
	/// </summary>
	public class MaildirWriter
	{
		String _path;
		String _curPath;
		long _lastUniqueId;

		private MaildirWriter(String path)
		{
			_path = path;
			_curPath = System.IO.Path.Combine(path, "cur");
			_lastUniqueId = new Random().Next();
		}

		public static MaildirWriter CreateMailDir(String path, String newFolderTitle) {
			//Creates a new maildir in an existing folder, and returns a new MaildirWriter object
			//to write messages to that maildir

			//Escape . chars in folder name, since . is the path delimiter for IMAP folders
			newFolderTitle = newFolderTitle.Replace(".", "_");

			//Prefix the dir name with a '.'
			String maildirPath = path + "." + newFolderTitle;

			if (Directory.Exists(maildirPath)) {
				Directory.Delete(maildirPath, true);
			}
			
			Directory.CreateDirectory(maildirPath);

			//Maildirs contain the mail in 'new' and 'cur', plus a 'tmp'.  
			//Since these are all existing messages, put everyting in 'cur'.
			Directory.CreateDirectory(System.IO.Path.Combine(maildirPath, "tmp"));
			Directory.CreateDirectory(System.IO.Path.Combine(maildirPath, "cur"));
			Directory.CreateDirectory(System.IO.Path.Combine(maildirPath, "new"));

			return new MaildirWriter(maildirPath);
		}

		public String Path {
			get {
				return _path;
			}
		}

		public void AddMessage(IMessage msg) {
			//Check the msg class.  If this is a non-delivery report, I don't know how to handle that
			String msgClass = MapiUtils.GetStringProperty(msg, Tags.PR_MESSAGE_CLASS);
			if (msgClass == "REPORT.IPM.Note.NDR") {
				return;
			}

			//A maildir is a folder full of files named thusly:
			// [timestamp].[uniqueid].[hostname]
			//
			// timestamp is simply the total seconds of the message time stamp
			// uniqueid is some system-generated ID that is unique within the namespace
			// of the hostname and timestamp.
			// hostname is the name of the host delivering the mail.
			//
			// since we're bulk-loading messages and can assume no other source is 
			// submitting them, cop out on the uniqueid and just use a randomly-seeded, increasing counter
			//
			// In addition, for msgs in 'cur' (which all of these are), the file name can be followed by the following:
			// (source: http://cr.yp.to/proto/maildir.html)

			//When you move a file from new to cur, you have to change its name from uniq to uniq:info. Make sure to preserve the uniq string, so that separate messages can't bump into each other.
			//
			//info is morally equivalent to the Status field used by mbox readers. It'd be useful to have MUAs agree on the meaning of info, so I'm keeping a list of info semantics. Here it is.
			//
			//info starting with "1,": Experimental semantics.
			//
			//info starting with "2,": Each character after the comma is an independent flag.
			//
			//    * Flag "P" (passed): the user has resent/forwarded/bounced this message to someone else.
			//    * Flag "R" (replied): the user has replied to this message.
			//    * Flag "S" (seen): the user has viewed this message, though perhaps he didn't read all the way through it.
			//    * Flag "T" (trashed): the user has moved this message to the trash; the trash will be emptied by a later user action.
			//    * Flag "D" (draft): the user considers this message a draft; toggled at user discretion.
			//    * Flag "F" (flagged): user-defined flag; toggled at user discretion. 
			//
			//New flags may be defined later. Flags must be stored in ASCII order: e.g., "2,FRS".
			//
			// From empirical study of Courier-IMAP Maildirs, there's also a 'S=whatever' before the flags.  So:
			// whatever:2,S=12345,RS. 
			//
			// Apparently this makes it faster for Courier to get messages sizes for quota enforcement purposes.  However
			//tbird at least don't seem to know what to do w/ this, so I won't include size info.
			//
			//The contents of each file are in standard RFC 2822 format.

			//TODO: Use PR_MSG_STATUS for R, D, T, F
			DateTime recvTime = MapiUtils.GetSysTimeProperty(msg, Tags.ptagMsgDeliveryTime);
			int size = MapiUtils.GetLongProperty(msg, Tags.PR_MESSAGE_SIZE);
			int flags = MapiUtils.GetLongProperty(msg, Tags.PR_MESSAGE_FLAGS);
			bool read = ((flags & (int)MAPI33.WellKnownValues.PR_MESSAGE_FLAGS.Read) != 0);
			bool draft = ((flags & (int)MAPI33.WellKnownValues.PR_MESSAGE_FLAGS.Unsent) != 0);

			//Update the unique id
			_lastUniqueId += 1 + new Random().Next(100);

			String name = recvTime.Ticks.ToString();
			name += ".";
			name += _lastUniqueId.ToString();
			name += ".";
			name += System.Net.Dns.GetHostName();
			//LAME: Colons in filenames aren't supported coz the path canonicalizer is too lazy to know the difference
			//between an alt data stream and just a plain colon.  Bogus.
			name += ";2,";
			//name += "S=" + size.ToString() + ","; tbird doesn't know what to do w/ this
			if (read) {
				name += "S";
			} 
			if (draft) {
				name += "D";
			}

			//Create a message object
			MailMessage outMsg = new MailMessage();
			outMsg.Date = recvTime;

			//Depending upon the message class, use different logic to populate the message
			if (msgClass == "IPM.Note") {
				ProcessMailMessage(msg, outMsg);
			} else if (msgClass == "IPM.Contact") {
				ProcessContact(msg, outMsg);
				return; //TODO: don't skip
			} else if (msgClass == "IPM.Calendar") {
				ProcessAppointment(msg, outMsg);
				return; //TODO: don't skip
			} else if (msgClass == "IPM.Task") {
				ProcessTask(msg, outMsg);
				return; //TODO: don't skip
			} else if (msgClass == "IPM.StickyNote") {
				ProcessNote(msg, outMsg);
				return; //TODO: don't skip
			} else {
				ProcessUnknownMsgClass(msg, outMsg);
				return; //TODO: don't skip
			}

			outMsg.AddCustomHeader("X-ConvertedFromMapi", 
				String.Format("Adam Nelson's Mapi to Maildir Converter.  {0}", DateTime.Now.ToString()));

			//Create the file in a temp location
			String tmpMsgFileName = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetTempFileName());
			outMsg.Save(tmpMsgFileName);

			//Read the msg from the temp file, and write it to the destination file,
			//replacing the Windows line endings (CR LF) with UNIX (LF)
			String msgFileName = System.IO.Path.Combine(_curPath, name);

			const String LF = "\xa";
			using (StreamReader sr = new StreamReader(tmpMsgFileName)) {
				using (StreamWriter sw = new StreamWriter(msgFileName)) {
					String line = null;
					while ( (line = sr.ReadLine()) != null) {
						sw.Write(line);
						sw.Write(LF);
					}
				}
			}

			File.Delete(tmpMsgFileName);

			//if there was a non-bogus receive time, set the date/time stamp
			//on the file to that time.
			if (recvTime != DateTime.MinValue) {
				//Set the date/time on the file to be the receive time
				//There seems to be an undocumented bug in the behavior of File.Set*Time.
				//Specifically, if it is used to set a time on a date with a different GMT
				//offset than the current date (eg, set the file to '3/1/04 12:00PM' (which is during
				//standard time in the Eastern time zone, and thus GMT-5) on '7/1/04' (which is
				//during daylight savings time in the Eastern time zone, and thus GMT-4), then
				//the resulting value displayed by Windows Explorer is one hour ahead.  To continue
				//the example above, explorer would display the time for the file as 
				//'1:00 PM' instead of '12:00 PM'.  It's as though the timezone translation
				//is being performed when it should not.
				//
				//There is newsgroup chatter suggesting this has been seen before, but not confirmed
				//by MS.  The Get*Time methods must reverse whatever transformation is applied
				//by the Set*Time methods, as they return the same time passed in.
				//
				//At any rate, the workaround is to apply a compensating reverse adjustment
				//That is, subtract from recvTime the difference in GMT offset on recvTime's date, 
				//and on the current date

				//For EST, this will be -300 minutes (-5 hours)
				//For EDT, -240 minutes (-4 hours)
				double thenGmtOffset = (recvTime - recvTime.ToUniversalTime()).TotalMinutes;
				double nowGmtOffset = (DateTime.Now - DateTime.Now.ToUniversalTime()).TotalMinutes;

				//To continue the example above, this will be an offset of -60 minutes
				double compensatingOffset = thenGmtOffset - nowGmtOffset;

				//Can't possibly be 24 hours or more.
				System.Diagnostics.Debug.Assert(compensatingOffset < 24*60);

				DateTime fsTime = recvTime.AddMinutes(compensatingOffset);

				File.SetCreationTime(msgFileName, fsTime);
				File.SetLastWriteTime(msgFileName, fsTime);
			}
		}

		public void AddFolder(IMAPIFolder folder) {
			//Add the messages in this folder to the maildir
			Error hr; 

			IMAPITable contentsTable;
			folder.GetContentsTable(0, out contentsTable);
			using(contentsTable) {
				MAPI33.MapiTypes.Value[,] rows;
				contentsTable.SetColumns(new Tags[] { Tags.PR_ENTRYID  }, IMAPITable.FLAGS.Default);

				for( ;contentsTable.QueryRows(1, 0, out rows) == Error.Success && rows.Length > 0; ) {
					//Get the message object for this entry id
					IUnknown unk;
					MAPI.TYPE objType;
					IMessage msg;

					hr = folder.OpenEntry(ENTRYID.Create(rows[0,0]), Guid.Empty, IMAPIContainer.FLAGS.BestAccess, out objType, out unk);
					if (hr != Error.Success) {
						throw new MapiException(hr);
					}

					msg = (IMessage)unk;
					unk.Dispose();

					using (msg) {
						AddMessage(msg);
					}
				}
			}
		}

		private void DumpMapiProperties(IMessage inMsg, MailMessage outMsg) {
			String header, body;
			//Get the original transport headers, and parse them out
			String xportHdrs = MapiUtils.GetStringProperty(inMsg, Tags.PR_TRANSPORT_MESSAGE_HEADERS);
			if (xportHdrs == null) {
				return;
			}

			//Each header line consists of the header, whitespace, colon, whitespace, value
			Regex hdrLineRex = new Regex(@"
(?# Match single- and multi-line SMTP headers )
(?<header>        (?# The header element... )
  [a-z0-9\-]+     (?# consists of letters, numbers, or hyphens.)
)                 (?# the header is followed by...)
\s*:\s*           (?# ...optional whitespace, a colon, additional optional whitespace...)
(?<value>         (?# ...and the value of the header, which is ...)
  .*?             (?# ...any character, possibly spanning multiple lines...)
)
(?=\r\n\S|\n\S|\z)(?# ...delimited by the start of another line w/ non-whitespace, or the end of the string)
", 
				RegexOptions.IgnoreCase | //Obviously, case-insensitive 
				RegexOptions.Multiline | //Need to match potentially multi-line SMTP headers 
				RegexOptions.Singleline | //and . matches newlines as well
				RegexOptions.IgnorePatternWhitespace //Ignore the pattern whitespace included for readability
				);

			foreach (Match match in hdrLineRex.Matches(xportHdrs)) {
				if (match.Success) {
					header = match.Groups["header"].Value;
					body = match.Groups["value"].Value;

					//If this header isn't one of the built-in ones that will be generated automatically
					//by OpenSmtp.net
					if (header.ToLower() != "to" &&
                        header.ToLower() != "from" &&
                        header.ToLower() != "reply-to" &&
                        header.ToLower() != "date" &&
                        header.ToLower() != "subject" &&
                        header.ToLower() != "cc" &&
                        header.ToLower() != "bcc" &&
                        header.ToLower() != "mime-version" &&
                        header.ToLower() != "content-type" &&
                        header.ToLower() != "content-transfer-encoding") {
						//OpenSmtp isn't smart enough to recognize multi-line header values, and wants to
						//quoted-printable them because of the CR/LF control chars.  So, straighten multi-line values out
						body = Regex.Replace(body, @"\r\n\s+|\n\s+", " ");
						outMsg.AddCustomHeader(header, body);
					}
				}
			}

			//Now dump all of the MAPI properties as X- headers just in case they're needed
			Tags[] propIds;
			inMsg.GetPropList(0, out propIds);
			foreach (Tags propId in propIds) {
				//Skip properties that are too big or redundant
				if (propId == Tags.ptagBody ||
					propId == Tags.ptagBodyHtml ||
					propId == Tags.ptagHtml ||
					propId == Tags.PR_BODY ||
					propId == Tags.PR_BODY_HTML ||
					propId == Tags.PR_RTF_COMPRESSED ||
					propId == Tags.PR_TRANSPORT_MESSAGE_HEADERS) {
					continue;
				}
				header = String.Format("X-{0}", propId);

				try {
					Value val = MapiUtils.GetProperty(inMsg, propId);

					if (val is MapiBinary) {
						//Binary values aren't good for much
						continue;
					}

					if (val == null) {
						body = "<null>";
					} else {
						body = val.ToString();
						//Cannot have line breaks in SMTP headers, so if there are any, escape them
						body = body.Replace("\n", @"\n");
						body = body.Replace("\r", @"\r");
						body = body.Replace("\t", @"\t");
					}

					outMsg.AddCustomHeader(header, body);
				} catch (MapiException e) {
					outMsg.AddCustomHeader("X-Exception", String.Format("Error getting property: {0}", e.Message));
				}
			}
		}

		private void ProcessMailMessage(IMessage msg, MailMessage outMsg) {
			//Copy some of the more important headers
			outMsg.From = MailMessage.CreateEmailAddress(MapiUtils.GetStringProperty(msg, Tags.PR_SENDER_EMAIL_ADDRESS),
				MapiUtils.GetStringProperty(msg, Tags.PR_SENDER_NAME));
			if (MapiUtils.GetStringProperty(msg, Tags.ptagSentRepresentingName) != null) {
                outMsg.ReplyTo = MailMessage.CreateEmailAddress(MapiUtils.GetStringProperty(msg, Tags.ptagSentRepresentingEmailAddr),
					MapiUtils.GetStringProperty(msg, Tags.ptagSentRepresentingName));
			}
			outMsg.Subject = MapiUtils.GetStringProperty(msg, Tags.PR_SUBJECT);

			IMAPITable recipientsTable = null;
			msg.GetRecipientTable(0, out recipientsTable);
			using (recipientsTable) {

				MAPI33.MapiTypes.Value[,] rows;
				recipientsTable.SetColumns(new Tags[] {Tags.PR_RECIPIENT_TYPE, Tags.PR_EMAIL_ADDRESS, Tags.PR_DISPLAY_NAME}, 
					IMAPITable.FLAGS.Default);

				for( ;recipientsTable.QueryRows(1, 0, out rows) == Error.Success && rows.Length > 0; ) {
					MAPI33.WellKnownValues.PR_RECIPIENT_TYPE recipType = (MAPI33.WellKnownValues.PR_RECIPIENT_TYPE)((MapiInt32)rows[0,0]).Value;
                    String emailAddr = null;
                    String emailName = null;

                    if (rows[0, 1] is MapiString) {
                        emailAddr = ((MapiString)rows[0, 1]).Value;
                    }

                    if (rows[0, 2] is MapiString) {
                        emailName = ((MapiString)rows[0, 2]).Value;
                    }

					AddressType type = AddressType.To;

					switch (recipType) {
						case MAPI33.WellKnownValues.PR_RECIPIENT_TYPE.To:
							type = AddressType.To;
							break;

						case MAPI33.WellKnownValues.PR_RECIPIENT_TYPE.Cc:
							type = AddressType.Cc;
							break;

						case MAPI33.WellKnownValues.PR_RECIPIENT_TYPE.Bcc:
							type = AddressType.Bcc;
							break;
					}

					outMsg.AddRecipient(emailAddr, emailName, type);
				}
			}

			//Set the body.  There can be a plain Body, which is the plain text version of the message, 
            //as well as an HtmlBody, which includes rich text formatting via HTML.
            //
            //This is complicated by the fact that, in MAPI, there are three possible sources
            //of the body: PR_BODY (plain text body), PR_BODY_HTML (HTML body), and PR_RTF_COMPRESSED 
            //(RTF body).  Based on empirical observation, there is nearly always a PR_BODY, which is not surprising
            //as every HTML mail client I've ever encountered sends multi-part MIME messages with a plain text version
            //as well as HTML.  PR_BODY_HTML *seems* to be the HTML version of the message as specified by the sender, 
            //and may not be present in plain text messages. 
            //
            //It only gets complicated with PR_RTF_COMPRESSED (or just PR_RTF for short).  Often, HTML messages
            //don't have a PR_HTML at all, but rather a PR_RTF that has a bunch of RTF tags wrapped around the original
            //HTML of the message.  All messages I've seen have a PR_RTF, however only messages which were originally
            //received as HTML and subsequently transmogrified into RTF include the \@fromhtml RTF tag.
            //
            //Thus, the existence of this \@fromhtml tag is a roundabout way of determining if a message came from
            //the sender as HTML or text (ignoring for now the third possiblity; an older Exchange client using RTF).
            //
            //Therefore, when setting BodyHtml, we first check if PR_RTF includes this \@fromhtml (that's what
            //MapiUtils.BodyRtfToHtml does, among other things).  If it doesn't, BodyHtml is set to null.  If it
            //does, PR_BODY_HTML is checked first, and used to populate BodyHtml if it's non-null.  Failing that, the 
            //unwrapped HTML from PR_RTF is used.  So:
            //
            // Body: Always PR_BODY
            // HtmlBody: PR_BODY_HTML if PR_BODY_HTML is non-null and PR_RTF includes the \@fromhtml tag
            //           The HTML embedded in PR_RTF if PR_BODY_HTML is null and PR_RTF includes \@fromhtml
            //           Else null.
            String propVal = null;
            if ((propVal = MapiUtils.GetLongStringProperty(msg, Tags.PR_BODY)) != null) {
                outMsg.Body = propVal;
			}

            //Check for HTML embedded in PR_RTF
			byte[] rtf = MapiUtils.GetBodyRtf(msg);
            if (rtf != null) {
                outMsg.HtmlBody = MapiUtils.BodyRtfToHtml(rtf);
                if (outMsg.HtmlBody != null) {
                    //There's HTML.  use PR_HTML if it's there, else stick w/ what we have
                    if ((propVal = MapiUtils.GetLongStringProperty(msg, Tags.PR_BODY_HTML)) != null) {
                        outMsg.HtmlBody = propVal;
                    }
                }
            }

			//Copy the attachments
			CopyAttachments(msg, outMsg);			

			//Output all of the MAPI properties as extended headers, just in case they are needed later
			DumpMapiProperties(msg, outMsg);
		}
		
		//TODO: Implement conversion of contacts to vCards and tasks to iCal files
		//http://www.cdolive.com/cdo10.htm has the necessary prop set IDs
		private void ProcessContact(IMessage msg, MailMessage outMsg) {
			outMsg.Body = "Conversion of contacts isn't implemented yet";
		}
		private void ProcessAppointment(IMessage msg, MailMessage outMsg) {
			outMsg.Body = "Conversion of appointments isn't implemented yet";
		}
		private void ProcessNote(IMessage msg, MailMessage outMsg) {
			outMsg.Body = "Conversion of tasks and notes isn't implemented yet";
		}
		private void ProcessTask(IMessage msg, MailMessage outMsg) {
			outMsg.Body = "Conversion of tasks and notes isn't implemented yet";
		}
		private void ProcessUnknownMsgClass(IMessage msg, MailMessage outMsg) {
			outMsg.Body = "Conversion of this message class isn't implemented yet";
		}

		private void CopyAttachments(IMessage msg, MailMessage outMsg) {
			MAPI33.IMAPITable attachmentsTable = null;
			msg.GetAttachmentTable(0, out attachmentsTable);
			using (attachmentsTable) {
				MAPI33.MapiTypes.Value[,] rows;
				attachmentsTable.SetColumns(new Tags[] {Tags.PR_ATTACH_FILENAME, Tags.PR_ATTACH_CONTENT_ID, Tags.PR_ATTACH_MIME_TAG, Tags.PR_ATTACH_ENCODING, Tags.PR_ATTACH_LONG_PATHNAME, Tags.PR_ATTACH_SIZE, Tags.PR_ATTACH_NUM}, 
					IMAPITable.FLAGS.Default);

				//Use OpenAttach to get an IAttachment

				int attachIdx = 0;
				for( ;attachmentsTable.QueryRows(1, 0, out rows) == Error.Success && rows.Length > 0; ) {
					try {
						attachIdx++;
						String attachFilename, attachMimeTag, attachContentId;
						uint attachNum;
						IAttachment attachObj;
						byte[] attachData;

						attachNum = (uint)((MapiInt32)rows[0,6]).Value;
						msg.OpenAttach(attachNum, IAttachment.IID, 0, out attachObj);
						using (attachObj) {
							if (rows[0,0] is MapiString) {
								attachFilename = ((MapiString)rows[0,0]).Value;
							} else {
								attachFilename = String.Format("Attach{0}.dat", attachIdx);
                            }

                            if (rows[0, 1] is MapiString) {
                                attachContentId = ((MapiString)rows[0, 1]).Value;
                            } else {
                                attachContentId = null;
                            }

							if (rows[0,2] is MapiString) {
								attachMimeTag = ((MapiString)rows[0,2]).Value;
							} else {
								attachMimeTag = String.Format("application/x-octet-stream");
							}

							IStream attachDataStream;
							IUnknown unk;
							Error hr = attachObj.OpenProperty(Tags.PR_ATTACH_DATA_BIN, IStream.IID, 0, 0, out unk);
							//OLE object attachments are PR_ATTACH_DATA_OBJ, so if BIN isn't found, try that
							if (hr == Error.NotFound) {
								hr = attachObj.OpenProperty(Tags.PR_ATTACH_DATA_OBJ, IStream.IID, 0, 0, out unk);
							}

							if (hr == Error.Success) {
								attachDataStream = (IStream)unk;
								unk.Dispose();
                                
								using (attachDataStream) {
									int bytesRead = 0;
									ArrayList streamContents = new ArrayList();
									byte[] buffer = new byte[1024];

									do {
										attachDataStream.Read(buffer, (uint)buffer.Length, out bytesRead);
										if (bytesRead == buffer.Length) {
											streamContents.AddRange(buffer);
										} else {
											for (int idx = 0; idx < (int)bytesRead; idx++) {
												streamContents.Add(buffer[idx]);
											}
										}
									} while (bytesRead > 0);

									attachData = (byte[])streamContents.ToArray(typeof(byte));
								}
							} else {
								attachData = System.Text.Encoding.ASCII.GetBytes(String.Format("Error getting contents of attachment '{0}' due to error '{1}'", attachFilename, hr));
							}

							//Use a MemoryStream to stream the attachment data into the Attachment object
							using (MemoryStream ms = new MemoryStream(attachData)) {
                                outMsg.AddAttachment(ms, attachFilename, attachMimeTag, attachContentId);
							} 
						}
					} catch (Exception e) {
						System.Diagnostics.Debug.WriteLine(e.ToString());
					}
				}
			}
		}
	}
}
