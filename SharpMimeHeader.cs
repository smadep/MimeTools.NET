// -----------------------------------------------------------------------
//
//   Copyright (C) 2003-2006 Angel Marin
// 
//   This file is part of SharpMimeTools
//
//   SharpMimeTools is free software; you can redistribute it and/or
//   modify it under the terms of the GNU Lesser General Public
//   License as published by the Free Software Foundation; either
//   version 2.1 of the License, or (at your option) any later version.
//
//   SharpMimeTools is distributed in the hope that it will be useful,
//   but WITHOUT ANY WARRANTY; without even the implied warranty of
//   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
//   Lesser General Public License for more details.
//
//   You should have received a copy of the GNU Lesser General Public
//   License along with SharpMimeTools; if not, write to the Free Software
//   Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
//
// -----------------------------------------------------------------------

using System;
using System.Collections.Specialized;
using System.Collections;
using System.Text;

namespace anmar.SharpMimeTools
{
    /// <summary>
    /// rfc 2822 header of a rfc 2045 entity
    /// </summary>
    public class SharpMimeHeader : IEnumerable
    {
#if LOG
		private static log4net.ILog log  = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
#endif
        private static Encoding default_encoding = Encoding.ASCII;
        private readonly SharpMimeMessageStream message;
        private readonly HybridDictionary headers;
        private String _cached_headers;
        private readonly long startpoint;
        private long endpoint;
        private long startbody;

        private struct HeaderInfo
        {
            public StringDictionary contenttype;
            public StringDictionary contentdisposition;
            public StringDictionary contentlocation;
            public MimeTopLevelMediaType TopLevelMediaType;
            public Encoding enc;
            public String subtype;

            public HeaderInfo(HybridDictionary headers)
            {
                TopLevelMediaType = new MimeTopLevelMediaType();
                enc = null;
                try
                {
                    contenttype = SharpMimeTools.parseHeaderFieldBody("Content-Type", headers["Content-Type"].ToString());
                    TopLevelMediaType = (MimeTopLevelMediaType)Enum.Parse(TopLevelMediaType.GetType(),
                                                                   contenttype["Content-Type"].Split('/')[0].Trim(),
                                                                   true);
                    subtype = contenttype["Content-Type"].Split('/')[1].Trim();
                    enc = SharpMimeTools.parseCharSet(contenttype["charset"]);
                }
                catch (Exception)
                {
                    enc = SharpMimeHeader.default_encoding;
                    contenttype = SharpMimeTools.parseHeaderFieldBody("Content-Type", String.Concat("text/plain; charset=", enc.BodyName));
                    TopLevelMediaType = MimeTopLevelMediaType.text;
                    subtype = "plain";
                }
                if (enc == null)
                {
                    enc = SharpMimeHeader.default_encoding;
                }
                // TODO: rework this
                try
                {
                    contentdisposition = SharpMimeTools.parseHeaderFieldBody("Content-Disposition", headers["Content-Disposition"].ToString());
                }
                catch (Exception)
                {
                    contentdisposition = new StringDictionary();
                }
                try
                {
                    contentlocation = SharpMimeTools.parseHeaderFieldBody("Content-Location", headers["Content-Location"].ToString());
                }
                catch (Exception)
                {
                    contentlocation = new StringDictionary();
                }
            }
        }
        private HeaderInfo mt;

        internal SharpMimeHeader(SharpMimeMessageStream message) : this(message, 0) { }

        internal SharpMimeHeader(SharpMimeMessageStream message, long startpoint)
        {
            this.startpoint = startpoint;
            this.message = message;
            if (this.startpoint == 0)
            {
                String line = this.message.ReadLine();
                // Perhaps there is part of the POP3 response
                if (line != null && line.Length > 3 && line[0] == '+' && line[1] == 'O' && line[2] == 'K')
                {
#if LOG
					if ( log.IsDebugEnabled ) log.Debug ("+OK present at top of the message");
#endif
                    this.startpoint = this.message.Position;
                }
                else
                    this.message.ReadLine_Undo(line);
            }
            headers = new HybridDictionary(2, true);
            parse();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpMimeHeader"/> class from a <see cref="System.IO.Stream"/>
        /// </summary>
        /// <param name="message"><see cref="System.IO.Stream"/> to read headers from</param>
        public SharpMimeHeader(System.IO.Stream message)
            : this(new SharpMimeMessageStream(message), 0)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        public SharpMimeHeader(Byte[] message)
            : this(new SharpMimeMessageStream(message), 0)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpMimeHeader"/> class from a <see cref="System.IO.Stream"/> starting from the specified point
        /// </summary>
        /// <param name="message">the <see cref="System.IO.Stream" /> to read headers from</param>
        /// <param name="startpoint">initial point of the <see cref="System.IO.Stream"/> where the headers start</param>
        public SharpMimeHeader(System.IO.Stream message, long startpoint)
            : this(new SharpMimeMessageStream(message), startpoint)
        {
        }
        
        /// <summary>
        /// Gets header fields
        /// </summary>
        /// <param name="name">field name</param>
        /// <remarks>Field names is case insentitive</remarks>
        public String this[Object name]
        {
            get
            {
                return getProperty(name.ToString());
            }
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public void Close()
        {
            _cached_headers = message.ReadLines(startpoint, endpoint);
            message.Close();
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public bool Contains(String name)
        {
            if (headers == null)
                parse();
            return headers.Contains(name);
        }
        
        /// <summary>
        /// Returns an enumerator that can iterate through the header fields
        /// </summary>
        /// <returns>A <see cref="System.Collections.IEnumerator" /> for the header fields</returns>
        public IEnumerator GetEnumerator()
        {
            return headers.GetEnumerator();
        }
        
        /// <summary>
        /// Returns the requested header field body.
        /// </summary>
        /// <param name="name">Header field name</param>
        /// <param name="defaultvalue">Value to return when the requested field is not present</param>
        /// <param name="uncomment"><b>true</b> to uncomment using <see cref="anmar.SharpMimeTools.SharpMimeTools.uncommentString" />; <b>false</b> to return the value unchanged.</param>
        /// <param name="rfc2047decode"><b>true</b> to decode <see cref="anmar.SharpMimeTools.SharpMimeTools.rfc2047decode" />; <b>false</b> to return the value unchanged.</param>
        /// <returns>Header field body</returns>
        public String GetHeaderField(String name, String defaultvalue, bool uncomment, bool rfc2047decode)
        {
            String tmp = getProperty(name);
            if (tmp == null)
                tmp = defaultvalue;
            else
            {
                if (uncomment)
                    tmp = SharpMimeTools.uncommentString(tmp);
                if (rfc2047decode)
                    tmp = SharpMimeTools.rfc2047decode(tmp);
            }
            return tmp;
        }
        
        private String getProperty(String name)
        {
            String Value;
            name = name.ToLower();
            parse();
            if (headers != null && headers.Count > 0 && name != null && name.Length > 0 && headers.Contains(name))
                Value = headers[name].ToString();
            else
                Value = null;
            return Value;
        }
        
        private bool parse()
        {
            bool error = false;
            if (headers.Count > 0)
            {
                return !error;
            }
            String line = String.Empty;
            message.SeekPoint(startpoint);
            message.Encoding = SharpMimeHeader.default_encoding;
            for (line = message.ReadUnfoldedLine(); line != null; line = message.ReadUnfoldedLine())
            {
                if (line.Length == 0)
                {
                    endpoint = message.Position_preRead;
                    startbody = message.Position;
                    message.ReadLine_Undo(line);
                    break;
                }
                else
                {
                    String[] headerline = line.Split(new Char[] { ':' }, 2);
                    if (headerline.Length == 2)
                    {
                        headerline[1] = headerline[1].TrimStart(new Char[] { ' ' });
                        if (headers.Contains(headerline[0]))
                        {
                            headers[headerline[0]] = String.Concat(headers[headerline[0]], "\r\n", headerline[1]);
                        }
                        else
                        {
                            headers.Add(headerline[0].ToLower(), headerline[1]);
                        }
                    }
                }
            }
            mt = new HeaderInfo(headers);
            return !error;
        }
        
        /// <summary>
        /// Gets the point where the headers end
        /// </summary>
        /// <value>Point where the headers end</value>
        public long BodyPosition
        {
            get
            {
                return startbody;
            }
        }
        
        /// <summary>
        /// Gets CC header field
        /// </summary>
        /// <value>CC header body</value>
        public String Cc
        {
            get { return GetHeaderField("Cc", String.Empty, true, false); }
        }
        
        /// <summary>
        /// Gets the number of header fields found
        /// </summary>
        public int Count
        {
            get
            {
                return headers.Count;
            }
        }
        
        /// <summary>
        /// Gets Content-Disposition header field
        /// </summary>
        /// <value>Content-Disposition header body</value>
        public String ContentDisposition
        {
            get { return GetHeaderField("Content-Disposition", String.Empty, true, false); }
        }
        
        /// <summary>
        /// Gets the elements found in the Content-Disposition header body
        /// </summary>
        /// <value><see cref="System.Collections.Specialized.StringDictionary"/> with the elements found in the header body</value>
        public StringDictionary ContentDispositionParameters
        {
            get
            {
                return mt.contentdisposition;
            }
        }
        
        /// <summary>
        /// Gets Content-ID header field
        /// </summary>
        /// <value>Content-ID header body</value>
        public String ContentID
        {
            get { return GetHeaderField("Content-ID", String.Empty, true, false); }
        }
        
        /// <summary>
        /// Gets Content-Location header field
        /// </summary>
        /// <value>Content-Location header body</value>
        public String ContentLocation
        {
            get { return GetHeaderField("Content-Location", String.Empty, true, false); }
        }
        
        /// <summary>
        /// Gets the elements found in the Content-Location header body
        /// </summary>
        /// <value><see cref="System.Collections.Specialized.StringDictionary"/> with the elements found in the header body</value>
        public StringDictionary ContentLocationParameters
        {
            get
            {
                return mt.contentlocation;
            }
        }
        
        /// <summary>
        /// Gets Content-Transfer-Encoding header field
        /// </summary>
        /// <value>Content-Transfer-Encoding header body</value>
        public String ContentTransferEncoding
        {
            get
            {
                String tmp = GetHeaderField("Content-Transfer-Encoding", null, false, false);
                if (tmp != null)
                {
                    tmp = tmp.ToLower();
                }
                return tmp;
            }
        }
        
        /// <summary>
        /// Gets Content-Type header field
        /// </summary>
        /// <value>Content-Type header body</value>
        public String ContentType
        {
            get { return GetHeaderField("Content-Type", String.Concat("text/plain; charset=", mt.enc.BodyName), false, false); }
        }
        
        /// <summary>
        /// Gets the elements found in the Content-Type header body
        /// </summary>
        /// <value><see cref="System.Collections.Specialized.StringDictionary"/> with the elements found in the header body</value>
        public StringDictionary ContentTypeParameters
        {
            get
            {
                return mt.contenttype;
            }
        }
        
        /// <summary>
        /// Gets Date header field
        /// </summary>
        /// <value>Date header body</value>
        public String Date
        {
            get { return GetHeaderField("Date", String.Empty, true, false); }
        }
        
        /// <summary>
        /// Gets <see cref="System.Text.Encoding"/> found on the headers and applies to the body
        /// </summary>
        /// <value><see cref="System.Text.Encoding"/> for the body</value>
        public Encoding Encoding
        {
            get
            {
                parse();
                return mt.enc;
            }
        }
        
        /// <summary>
        /// Gets or sets the default <see cref="System.Text.Encoding" /> used when it isn't defined otherwise.
        /// </summary>
        /// <value>The <see cref="System.Text.Encoding" /> used when it isn't defined otherwise</value>
        /// <remarks>The default value is <see cref="System.Text.ASCIIEncoding" /> as defined in RFC 2045 section 5.2.<br />
        /// If you change this value you'll be violating this rfc section.</remarks>
        public static Encoding EncodingDefault
        {
            get { return default_encoding; }
            set
            {
                if (value != null && !value.BodyName.Equals(String.Empty))
                    default_encoding = value;
                else
                    default_encoding = Encoding.ASCII;
            }
        }
        
        /// <summary>
        /// Gets From header field
        /// </summary>
        /// <value>From header body</value>
        public String From
        {
            get { return GetHeaderField("From", String.Empty, true, false); }
        }
        
        /// <summary>
        /// Gets Raw headers
        /// </summary>
        /// <value>From header body</value>
        public String RawHeaders
        {
            get
            {
                if (_cached_headers != null)
                    return _cached_headers;
                else
                    return message.ReadLines(startpoint, endpoint);
            }
        }
        
        /// <summary>
        /// Gets Message-ID header field
        /// </summary>
        /// <value>Message-ID header body</value>
        public String MessageID
        {
            get { return GetHeaderField("Message-ID", String.Empty, true, false); }
        }
        
        /// <summary>
        /// Gets reply address as defined by <c>rfc 2822</c>
        /// </summary>
        /// <value>Reply address</value>
        public String Reply
        {
            get
            {
                if (!ReplyTo.Equals(String.Empty))
                    return ReplyTo;
                else
                    return From;
            }
        }
        
        /// <summary>
        /// Gets Reply-To header field
        /// </summary>
        /// <value>Reply-To header body</value>
        public String ReplyTo
        {
            get { return GetHeaderField("Reply-To", String.Empty, true, false); }
        }
        
        /// <summary>
        /// Gets Return-Path header field
        /// </summary>
        /// <value>Return-Path header body</value>
        public String ReturnPath
        {
            get { return GetHeaderField("Return-Path", String.Empty, true, false); }
        }
        
        /// <summary>
        /// Gets Sender header field
        /// </summary>
        /// <value>Sender header body</value>
        public String Sender
        {
            get { return GetHeaderField("Sender", String.Empty, true, false); }
        }
        
        /// <summary>
        /// Gets Subject header field
        /// </summary>
        /// <value>Subject header body</value>
        public String Subject
        {
            get { return GetHeaderField("Subject", String.Empty, false, false); }
        }
        
        /// <summary>
        /// Gets SubType from Content-Type header field
        /// </summary>
        /// <value>SubType from Content-Type header field</value>
        public String SubType
        {
            get
            {
                parse();
                return mt.subtype;
            }
        }
        
        /// <summary>
        /// Gets To header field
        /// </summary>
        /// <value>To header body</value>
        public String To
        {
            get { return GetHeaderField("To", String.Empty, true, false); }
        }
        
        /// <summary>
        /// Gets top-level media type from Content-Type header field
        /// </summary>
        /// <value>Top-level media type from Content-Type header field</value>
        public MimeTopLevelMediaType TopLevelMediaType
        {
            get
            {
                parse();
                return mt.TopLevelMediaType;
            }
        }
        
        /// <summary>
        /// Gets Version header field
        /// </summary>
        /// <value>Version header body</value>
        public String Version
        {
            get { return GetHeaderField("Version", "1.0", true, false); }
        }
    }
}
