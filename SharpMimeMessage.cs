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
using System.Collections;
using System.IO;
using System.Diagnostics;

namespace anmar.SharpMimeTools
{
	/// <summary>
	/// rfc 2045 entity
	/// </summary>
	public class SharpMimeMessage : IEnumerable, IDisposable {
		private struct MessageInfo {
			internal long start;
			internal long start_body;
			internal long end;
			internal SharpMimeHeader header;
			internal SharpMimeMessageCollection parts;
            internal MessageInfo(SharpMimeMessageStream m, long start)
            {
                this.start = start;
                header = new SharpMimeHeader(m, this.start);
                start_body = header.BodyPosition;
                end = -1;
                parts = new SharpMimeMessageCollection();
            }
		}

        private readonly SharpMimeMessageStream message;
        private MessageInfo mi;

		/// <summary>
		/// Initializes a new instance of the <see cref="SharpMimeMessage"/> class from a <see cref="System.IO.Stream"/>
		/// </summary>
		/// <param name="message"><see cref="System.IO.Stream" /> to read the message from</param>
        public SharpMimeMessage(Stream message)
        {
            this.message = new SharpMimeMessageStream(message);
            mi = new MessageInfo(this.message, this.message.Stream.Position);
        }

        /// <summary>
        /// Gets header fields for this entity
        /// </summary>
        /// <param name="name">field name</param>
        /// <remarks>Field names is case insentitive</remarks>
        public String this[Object name]
        {
            get { return mi.header[name.ToString()]; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public String Body
        {
            get
            {
                return GetRawBody(false);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public String BodyDecoded
        {
            get
            {
                switch (Header.ContentTransferEncoding)
                {
                    case "quoted-printable":
                        return SharpMimeTools.QuotedPrintable2Unicode(mi.header.Encoding, GetRawBody(false));
                    case "base64":
                        Byte[] tmp = null;
                        try
                        {
                            tmp = Convert.FromBase64String(GetRawBody(false));
                        }
                        catch (Exception e)
                        {
                            Trace.Fail(e.Message, e.StackTrace);
                        }
                        if (tmp != null)
                            return mi.header.Encoding.GetString(tmp);
                        else
                            return String.Empty;
                    default:
                        break;
                }
                return GetRawBody(false);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public String Disposition
        {
            get
            {
                return Header.ContentDispositionParameters["Content-Disposition"];
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public SharpMimeHeader Header
        {
            get
            {
                return mi.header;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool IsBrowserDisplay
        {
            get
            {
                switch (mi.header.TopLevelMediaType)
                {
                    case MimeTopLevelMediaType.audio:
                    case MimeTopLevelMediaType.image:
                    case MimeTopLevelMediaType.text:
                    case MimeTopLevelMediaType.video:
                        return true;
                    case MimeTopLevelMediaType.application:
                    case MimeTopLevelMediaType.message:
                    case MimeTopLevelMediaType.multipart:
                    default:
                        return false;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool IsMultipart
        {
            get
            {
                switch (mi.header.TopLevelMediaType)
                {
                    case MimeTopLevelMediaType.multipart:
                    case MimeTopLevelMediaType.message:
                        return true;
                    case MimeTopLevelMediaType.application:
                    case MimeTopLevelMediaType.audio:
                    case MimeTopLevelMediaType.image:
                    case MimeTopLevelMediaType.text:
                    case MimeTopLevelMediaType.video:
                    default:
                        return false;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool IsTextBrowserDisplay
        {
            get
            {
                if (mi.header.TopLevelMediaType.Equals(MimeTopLevelMediaType.text) && mi.header.SubType.Equals("plain"))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public String Name
        {
            get
            {
                parse();
                String param = Header.ContentDispositionParameters["filename"];

                if (param == null)
                {
                    param = Header.ContentTypeParameters["name"];
                }
                if (param == null)
                {
                    param = Header.ContentLocationParameters["Content-Location"];
                }
                return SharpMimeTools.GetFileName(param);
            }
        }

        internal SharpMimeMessageCollection Parts
        {
            get
            {
                parse();
                return mi.parts;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public int PartsCount
        {
            get
            {
                parse();
                return mi.parts.Count;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public long Size
        {
            get
            {
                parse();
                return mi.end - mi.start_body;
            }
        }
		
        /// <summary>
		/// Clears the parts references contained in this instance and calls the <b>Close</b> method in those parts.
		/// </summary>
		/// <remarks>This method does not close the underling <see cref="System.IO.Stream" /> used to create this instance.</remarks>
        public void Close()
        {
            foreach (SharpMimeMessage part in mi.parts)
                part.Close();
            mi.parts.Clear();
        }
		
        /// <summary>
		/// Dumps the body of this entity into a <see cref="System.IO.Stream"/>
		/// </summary>
		/// <param name="stream"><see cref="System.IO.Stream" /> where we want to write the body</param>
		/// <returns><b>true</b> OK;<b>false</b> if write operation fails</returns>
        public bool DumpBody(Stream stream)
        {
            if (stream == null)
                return false;
            bool error = false;
            bool raw = false;
            if (stream.CanWrite)
            {
                Byte[] buffer = null;
                switch (Header.ContentTransferEncoding)
                {
                    case "quoted-printable":
                        buffer = mi.header.Encoding.GetBytes(BodyDecoded);
                        break;
                    case "base64":
                        try
                        {
                            buffer = Convert.FromBase64String(GetRawBody(true));
                        }
                        catch (Exception)
                        {
                            error = true;
                        }
                        break;
                    case "7bit":
                    case "8bit":
                    case "binary":
                    case null:
                        raw = true;
                        break;
                    default:
                        error = true;
                        break;
                }
                try
                {
                    if (!error)
                    {
                        if (raw)
                        {
                            message.SaveTo(stream, mi.start_body, mi.end);
                        }
                        else if (buffer != null)
                        {
                            stream.Write(buffer, 0, buffer.Length);
                        }
                    }
                }
                catch (Exception)
                {
                    error = true;
                }
                buffer = null;
            }
            else
            {
                error = true;
            }
            return !error;
        }
		
        /// <summary>
		/// Dumps the body of this entity into a file
		/// </summary>
		/// <param name="path">path of the destination folder</param>
		/// <returns><see cref="System.IO.FileInfo" /> that represents the file where the body has been saved</returns>
        public FileInfo DumpBody(String path)
        {
            return DumpBody(path, Name);
        }
		
        /// <summary>
		/// Dumps the body of this entity into a file
		/// </summary>
		/// <param name="path">path of the destination folder</param>
		/// <param name="generatename">true if the filename must be generated incase we can't find a valid one in the headers</param>
		/// <returns><see cref="System.IO.FileInfo" /> that represents the file where the body has been saved</returns>
        public FileInfo DumpBody(String path, bool generatename)
        {
            String name = Name;
            if (name == null && generatename)
                name = String.Format("generated_{0}.{1}", GetHashCode(), Header.SubType);
            return DumpBody(path, name);
        }
		
        /// <summary>
		/// Dumps the body of this entity into a file
		/// </summary>
		/// <param name="path">path of the destination folder</param>
		/// <param name="name">name of the file</param>
		/// <returns><see cref="System.IO.FileInfo" /> that represents the file where the body has been saved</returns>
        public FileInfo DumpBody(String path, String name)
        {
            FileInfo file = null;
            if (name != null)
            {
                name = Path.GetFileName(name);
                // Dump file contents
                try
                {
                    DirectoryInfo dir = new DirectoryInfo(path);
                    dir.Create();
                    try
                    {
                        file = new FileInfo(Path.Combine(path, name));
                    }
                    catch (ArgumentException)
                    {
                        file = null;
                    }
                    if (file != null && dir.Exists)
                    {
                        if (dir.FullName.Equals(new DirectoryInfo(file.Directory.FullName).FullName))
                        {
                            if (!file.Exists)
                            {
                                Stream stream = null;
                                try
                                {
                                    stream = file.Create();
                                }
                                catch (Exception e)
                                {
                                    Trace.Fail(e.Message, e.StackTrace);
                                }
                                bool error = !DumpBody(stream);
                                if (stream != null)
                                    stream.Close();
                                stream = null;
                                if (error)
                                {
                                    if (stream != null)
                                        file.Delete();
                                }
                                else
                                {
                                    // The file should be there
                                    file.Refresh();
                                    // Set file dates
                                    if (Header.ContentDispositionParameters.ContainsKey("creation-date"))
                                        file.CreationTime = SharpMimeTools.parseDate(Header.ContentDispositionParameters["creation-date"]);
                                    if (Header.ContentDispositionParameters.ContainsKey("modification-date"))
                                        file.LastWriteTime = SharpMimeTools.parseDate(Header.ContentDispositionParameters["modification-date"]);
                                    if (Header.ContentDispositionParameters.ContainsKey("read-date"))
                                        file.LastAccessTime = SharpMimeTools.parseDate(Header.ContentDispositionParameters["read-date"]);
                                }
                            }
                        }
                    }
                    dir = null;
                }
                catch (Exception)
                {
                    try
                    {
                        if (file != null)
                        {
                            file.Refresh();
                            if (file.Exists)
                                file.Delete();
                        }
                    }
                    catch (Exception e)
                    {
                        Trace.Fail(e.Message, e.StackTrace);
                    }
                    file = null;
                }
            }
            return file;
        }
		
        /// <summary>
		/// Returns an enumerator that can iterate through the parts of a multipart entity
		/// </summary>
		/// <returns>A <see cref="System.Collections.IEnumerator" /> for the parts of a multipart entity</returns>
        public IEnumerator GetEnumerator()
        {
            parse();
            return mi.parts.GetEnumerator();
        }
		
        /// <summary>
		/// Returns the requested part of a multipart entity
		/// </summary>
		/// <param name="index">index of the requested part</param>
		/// <returns>A <see cref="anmar.SharpMimeTools.SharpMimeMessage" /> for the requested part</returns>
        public SharpMimeMessage GetPart(int index)
        {
            return Parts[index];
        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (message != null)
                {
                    message.Dispose();
                }
            }
        }

        private SharpMimeMessage(SharpMimeMessageStream message, long startpoint)
        {
            this.message = message;
            mi = new MessageInfo(this.message, startpoint);
        }

        private SharpMimeMessage(SharpMimeMessageStream message, long startpoint, long endpoint)
        {
            this.message = message;
            mi = new MessageInfo(this.message, startpoint) { end = endpoint };
        }

        private String GetRawBody(bool rawparts)
        {
            parse();
            if (rawparts || mi.parts.Count == 0)
            {
                message.Encoding = mi.header.Encoding;
                if (mi.end == -1)
                {
                    return message.ReadAll(mi.start_body);
                }
                else
                {
                    return message.ReadLines(mi.start_body, mi.end);
                }
            }
            else
            {
                return null;
            }
        }
        
        private bool parse()
        {
            
            if (!IsMultipart || Equals(mi.parts.Parent))
            {
                return true;
            }
            switch (mi.header.TopLevelMediaType)
            {
                case MimeTopLevelMediaType.message:
                    mi.parts.Parent = this;
                    SharpMimeMessage message = new SharpMimeMessage(this.message, mi.start_body, mi.end);
                    mi.parts.Add(message);
                    break;
                case MimeTopLevelMediaType.multipart:
                    this.message.SeekPoint(mi.start_body);
                    String line;
                    mi.parts.Parent = this;
                    String boundary_start = String.Concat("--", mi.header.ContentTypeParameters["boundary"]);
                    String boundary_end = String.Concat("--", mi.header.ContentTypeParameters["boundary"], "--");
                    for (line = this.message.ReadLine(); line != null; line = this.message.ReadLine())
                    {
                        // It can't be a boundary line
                        if (line.Length < 3)
                            continue;
                        // Match start boundary line
                        if (line.Length == boundary_start.Length && line == boundary_start)
                        {
                            if (mi.parts.Count > 0)
                            {
                                mi.parts[mi.parts.Count - 1].mi.end = this.message.Position_preRead;
                            }
                            SharpMimeMessage msg = new SharpMimeMessage(this.message, this.message.Position);
                            mi.parts.Add(msg);
                            // Match end boundary line
                        }
                        else if (line.Length == boundary_end.Length && line == boundary_end)
                        {
                            mi.end = this.message.Position_preRead;
                            if (mi.parts.Count > 0)
                            {
                                mi.parts[mi.parts.Count - 1].mi.end = this.message.Position_preRead;
                            }
                            break;
                        }
                    }
                    break;
                case MimeTopLevelMediaType.audio:
                case MimeTopLevelMediaType.image:
                case MimeTopLevelMediaType.text:
                case MimeTopLevelMediaType.video:
                case MimeTopLevelMediaType.application:
                default:
                    break;
            }
            return !false;
        }
    }
}
