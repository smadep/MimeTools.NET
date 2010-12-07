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
//   Foundation, Inc., 51 Franklin Street, Fifth Floor,
//   Boston, MA  02110-1301  USA
//
// -----------------------------------------------------------------------

using System;
using System.Collections;
using System.IO;
using System.Collections.Generic;

namespace anmar.SharpMimeTools
{
    /// <summary>
    /// Decodes <a href="http://msdn.microsoft.com/library/en-us/mapi/html/ca148ec3-8586-4c74-8ff8-cd542256e385.asp">ms-tnef</a> streams (those winmail.dat attachments). 
    /// </summary>
    /// <remarks>Only tnef attributes related to attachments are decoded right now. All the MAPI properties encoded in the stream (rtf body, ...) are ignored. <br />
    /// While decoding, the cheksum of each attribute is verified to ensure the tnef stream is not corupt.</remarks>
    public class SharpTnefMessage : IDisposable
    {
        private enum TnefLvlType : byte
        {
            Message = 0x01,
            Attachment = 0x02,
            Unknown = Byte.MaxValue
        }

        // TNEF attributes
        private enum TnefAttribute : ushort
        {
            Owner = 0x0000,
            SentFor = 0x0001,
            Delegate = 0x0002,
            DateStart = 0x0006,
            DateEnd = 0x0007,
            AIDOwner = 0x0008,
            Requestres = 0x0009,
            From = 0x8000,
            Subject = 0x8004,
            DateSent = 0x8005,
            DateRecd = 0x8006,
            MessageStatus = 0x8007,
            MessageClass = 0x8008,
            MessageId = 0x8009,
            ParentId = 0x800a,
            ConversationId = 0x800b,
            Body = 0x800c,
            Priority = 0x800d,
            AttachData = 0x800f,
            AttachTitle = 0x8010,
            AttachMetafile = 0x8011,
            AttachCreateDate = 0x8012,
            AttachModifyDate = 0x8013,
            DateModify = 0x8020,
            AttachTransportFilename = 0x9001,
            AttachRendData = 0x9002,
            MapiProps = 0x9003,
            RecipTable = 0x9004,
            Attachment = 0x9005,
            TnefVersion = 0x9006,
            OEMCodepage = 0x9007,
            OriginalMessageClass = 0x9008,
            Unknown = UInt16.MaxValue
        }

        // TNEF data types
        private enum TnefDataType : ushort
        {
            atpTriples = 0x0000,
            atpString = 0x0001,
            atpText = 0x0002,
            atpDate = 0x0003,
            atpShort = 0x0004,
            atpLong = 0x0005,
            atpByte = 0x0006,
            atpWord = 0x0007,
            atpDword = 0x0008,
            atpMax = 0x0009,
            Unknown = UInt16.MaxValue
        }

        // TNEF signature
        private const int TnefSignature = 0x223e9f78;
        private BinaryReader _reader;
        private String _body;

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpTnefMessage" /> class based on the supplied <see cref="System.IO.Stream" />.
        /// </summary>
        /// <param name="input"><see cref="System.IO.Stream" /> that contains the ms-tnef strream.</param>
        /// <remarks>The tnef stream isn't automatically parsed, you must call <see cref="Parse()" /> or <see cref="Parse(string)" />.</remarks>
        public SharpTnefMessage(Stream input)
        {
            if (input != null && input.CanRead)
            {
                if (input is BufferedStream)
                    _reader = new BinaryReader(input);
                else
                    _reader = new BinaryReader(new BufferedStream(input));
            }
        }

        /// <summary>
        /// Gets a list that contains the attachments found in the tnef stream.
        /// </summary>
        /// <value>A list that contains the <see cref="SharpAttachment" /> found in the tnef stream. The <b>null</b> reference is retuned when no attachments found.</value>
        /// <remarks>Each attachment is a <see cref="SharpAttachment" /> instance.</remarks>
        public List<SharpAttachment> Attachments { get; private set; }
        
        /// <summary>
        /// Gets a the text body from the ms-tnef stream (<b>BODY</b> tnef attribute).
        /// </summary>
        /// <value>Text body from the ms-tnef stream (<b>BODY</b> tnef attribute). Or the <b>null</b> reference if the attribute is not part of the stream.</value>
        public String TextBody
        {
            get { return _body; }
        }
        
        /// <summary>
        /// Closes and releases the reading resources associated with this instance. 
        /// </summary>
        /// <remarks>Be carefull before calling this method, as it also close the underlying <see cref="System.IO.Stream" />.</remarks>
        public void Close()
        {
            if (_reader != null)
                _reader.Close();
            _reader = null;
        }
        
        /// <summary>
        /// Parses the ms-tnef stream from the provided <see cref="System.IO.Stream" />.
        /// </summary>
        /// <returns><b>true</b> if parsing is successful. <b>false</b> otherwise.</returns>
        /// <remarks>The attachments found are saved in memory as <see cref="System.IO.MemoryStream" /> instances.</remarks>
        public bool Parse()
        {
            return Parse(null);
        }
        
        /// <summary>
        /// Parses the ms-tnef stream from the provided <see cref="System.IO.Stream" />.
        /// </summary>
        /// <param name="path">A <see cref="System.String" /> specifying the path on which to save the attachments found. Specify the <b>null</b> reference to save them in memory as  <see cref="System.IO.MemoryStream" /> instances instead.</param>
        /// <returns><b>true</b> if parsing is successful. <b>false</b> otherwise.</returns>
        public bool Parse(String path)
        {
            if (_reader == null || !_reader.BaseStream.CanRead)
                return false;
            int sig = ReadInt();
            if (sig != TnefSignature)
            {
                return false;
            }
            bool error = false;
            Attachments = new List<SharpAttachment>();
            ushort key = ReadUInt16();
            System.Text.Encoding enc = SharpMimeHeader.EncodingDefault;
            SharpAttachment attachment_cur = null;
            for (Byte cur = ReadByte(); cur != Byte.MinValue; cur = ReadByte())
            {
                TnefLvlType lvl = (TnefLvlType)SharpMimeTools.ParseEnum(typeof(TnefLvlType), cur, TnefLvlType.Unknown);
                // Type
                int type = ReadInt();
                // Size
                int size = ReadInt();
                // Attibute name and type
                TnefAttribute att_n = (TnefAttribute)SharpMimeTools.ParseEnum(typeof(TnefAttribute), (ushort)((type << 16) >> 16), TnefAttribute.Unknown);
                TnefDataType att_t = (TnefDataType)SharpMimeTools.ParseEnum(typeof(TnefDataType), (ushort)(type >> 16), TnefDataType.Unknown);
                if (lvl == TnefLvlType.Unknown || att_n == TnefAttribute.Unknown || att_t == TnefDataType.Unknown)
                {
                    error = true;
                    break;
                }
                // Read data
                Byte[] buffer = ReadBytes(size);
                // Read checkSum
                ushort checksum = ReadUInt16();
                // Verify checksum
                if (!VerifyChecksum(buffer, checksum))
                {
                    error = true;
                    break;
                }
                size = buffer.Length;
                if (lvl == TnefLvlType.Message)
                {
                    // Text body
                    if (att_n == TnefAttribute.Body)
                    {
                        if (att_t == TnefDataType.atpString)
                        {
                            _body = enc.GetString(buffer, 0, size);
                        }
                        // Message mapi props (html body, rtf body, ...)
                    }
                    else if (att_n == TnefAttribute.MapiProps)
                    {
                        ReadMapi(buffer);
                        // Stream Codepage
                    }
                    else if (att_n == TnefAttribute.OEMCodepage)
                    {
                        uint codepage1 = (uint)(buffer[0] + (buffer[1] << 8) + (buffer[2] << 16) + (buffer[3] << 24));
                        if (codepage1 > 0)
                        {
                            try
                            {
                                enc = System.Text.Encoding.GetEncoding((int)codepage1);
                            }
                            catch (Exception) { }
                        }
                    }
                }
                else if (lvl == TnefLvlType.Attachment)
                {
                    // Attachment start
                    if (att_n == TnefAttribute.AttachRendData)
                    {
                        String name = String.Concat("generated_", key, "_", (Attachments.Count + 1), ".binary");
                        if (path == null)
                        {
                            attachment_cur = new SharpAttachment();
                        }
                        else
                        {
                            attachment_cur = new SharpAttachment(new FileInfo(Path.Combine(path, name)));
                        }

                        // Attachment name
                        attachment_cur.Name = name;
                    }
                    else if (att_n == TnefAttribute.AttachTitle)
                    {
                        if (attachment_cur != null && att_t == TnefDataType.atpString && buffer != null)
                        {
                            // NULL terminated string
                            if (buffer[size - 1] == '\0')
                            {
                                size--;
                            }
                            if (size > 0)
                            {
                                String name = enc.GetString(buffer, 0, size);
                                if (name.Length > 0)
                                {
                                    attachment_cur.Name = name;
                                    // Content already saved, so we have to rename
                                    if (attachment_cur.SavedFile != null && attachment_cur.SavedFile.Exists)
                                    {
                                        try
                                        {
                                            attachment_cur.SavedFile.MoveTo(Path.Combine(path, attachment_cur.Name));
                                        }
                                        catch (Exception) { }
                                    }
                                }
                            }
                        }
                        // Modification and creation date
                    }
                    else if (att_n == TnefAttribute.AttachModifyDate || att_n == TnefAttribute.AttachCreateDate)
                    {
                        if (attachment_cur != null && att_t == TnefDataType.atpDate && buffer != null && size == 14)
                        {
                            int pos = 0;
                            DateTime date = new DateTime((buffer[pos++] + (buffer[pos++] << 8)), (buffer[pos++] + (buffer[pos++] << 8)), (buffer[pos++] + (buffer[pos++] << 8)), (buffer[pos++] + (buffer[pos++] << 8)), (buffer[pos++] + (buffer[pos++] << 8)), (buffer[pos++] + (buffer[pos++] << 8)));
                            if (att_n == TnefAttribute.AttachModifyDate)
                            {
                                attachment_cur.LastWriteTime = date;
                            }
                            else if (att_n == TnefAttribute.AttachCreateDate)
                            {
                                attachment_cur.CreationTime = date;
                            }
                        }
                        // Attachment data
                    }
                    else if (att_n == TnefAttribute.AttachData)
                    {
                        if (attachment_cur != null && att_t == TnefDataType.atpByte && buffer != null)
                        {
                            if (attachment_cur.SavedFile != null)
                            {
                                FileStream stream = null;
                                try
                                {
                                    stream = attachment_cur.SavedFile.OpenWrite();
                                }
                                catch (Exception e)
                                {
                                    error = true;
                                    break;
                                }
                                stream.Write(buffer, 0, size);
                                stream.Flush();
                                attachment_cur.Size = stream.Length;
                                stream.Close();
                                stream = null;
                                attachment_cur.SavedFile.Refresh();
                                // Is name has changed, we have to rename
                                if (attachment_cur.SavedFile.Name != attachment_cur.Name)
                                    try
                                    {
                                        attachment_cur.SavedFile.MoveTo(Path.Combine(path, attachment_cur.Name));
                                    }
                                    catch (Exception) { }
                            }
                            else
                            {
                                attachment_cur.Stream.Write(buffer, 0, size);
                                attachment_cur.Stream.Flush();
                                if (attachment_cur.Stream.CanSeek)
                                    attachment_cur.Stream.Seek(0, SeekOrigin.Begin);
                                attachment_cur.Size = attachment_cur.Stream.Length;
                            }
                            Attachments.Add(attachment_cur);
                        }
                        // Attachment mapi props
                    }
                    else if (att_n == TnefAttribute.Attachment)
                    {
                        ReadMapi(buffer);
                    }
                }
            }
            if (Attachments.Count == 0)
                Attachments = null;
            return !error;
        }

        private static void ReadMapi(Byte[] data)
        {
            int pos = 0;
            ushort count = (ushort)(data[pos++] + (data[pos++] << 8));
            if (count == 0)
                return;
            //FIXME: Read each mapi prop
        }

        private static bool VerifyChecksum(Byte[] data, ushort checksum)
        {
            if (data == null)
                return false;
            ushort checksum_calc = 0;
            for (int i = 0, count = data.Length; i < count; i++)
            {
                checksum_calc += data[i];
            }
            checksum_calc = (ushort)(checksum_calc % 65536);
            return (checksum_calc == checksum);
        }

        private Byte ReadByte()
        {
            Byte cur;
            try
            {
                cur = _reader.ReadByte();
            }
            catch (Exception)
            {
                cur = Byte.MinValue;
            }
            return cur;
        }
        
        private Byte[] ReadBytes(int length)
        {
            if (length <= 0)
                return null;
            Byte[] buffer = null;
            try
            {
                buffer = _reader.ReadBytes(length);
            }
            catch (Exception) { }
            return buffer;
        }
        
        private int ReadInt()
        {
            int cur;
            try
            {
                cur = _reader.ReadInt32();
            }
            catch (Exception)
            {
                cur = Int32.MinValue;
            }
            return cur;
        }
        
        private ushort ReadUInt16()
        {
            ushort cur;
            try
            {
                cur = _reader.ReadUInt16();
            }
            catch (Exception)
            {
                cur = UInt16.MinValue;
            }
            return cur;
        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_reader != null)
                {
                    _reader.Close();
                }
            }
        }
    }
}
