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
using System.Web;
using System.Collections.Generic;

namespace anmar.SharpMimeTools
{
    /// <summary>
    /// This class provides a simplified way of parsing messages. 
    /// </summary>
    /// <remarks> All the mime complexity is handled internally and all the content is exposed
    /// parsed and decoded. The code takes care of comments, RFC 2047, encodings, etc.</remarks>
    /// <example>Parse a message read from a file enabling the uuencode and ms-tnef decoding and saving attachments to disk.
    /// <code>
    /// System.IO.FileStream msg = new System.IO.FileStream("message_file.txt", System.IO.FileMode.Open);
    /// anmar.SharpMimeTools.SharpMessage message = new anmar.SharpMimeTools.SharpMessage(msg, SharpDecodeOptions.Default|SharpDecodeOptions.DecodeTnef|SharpDecodeOptions.UuDecode);
    /// msg.Close();
    /// Console.WriteLine(System.String.Concat("From:    [", message.From, "][", message.FromAddress, "]"));
    /// Console.WriteLine(System.String.Concat("To:      [", message.To, "]"));
    /// Console.WriteLine(System.String.Concat("Subject: [", message.Subject, "]"));
    /// Console.WriteLine(System.String.Concat("Date:    [", message.Date, "]"));
    /// Console.WriteLine(System.String.Concat("Body:    [", message.Body, "]"));
    /// if ( message.Attachments!=null ) {
    /// 	foreach ( anmar.SharpMimeTools.SharpAttachment attachment in message.Attachments ) {
    /// 		attachment.Save(System.Environment.CurrentDirectory, false);
    /// 		Console.WriteLine(System.String.Concat("Attachment: [", attachment.SavedFile.FullName, "]"));
    /// 	}
    /// }
    /// </code>
    /// </example>
    /// <example>This sample shows how simple is to parse an e-mail message read from a file.
    /// <code>
    /// System.IO.FileStream msg = new System.IO.FileStream("message_file.txt", System.IO.FileMode.Open);
    /// anmar.SharpMimeTools.SharpMessage message = new anmar.SharpMimeTools.SharpMessage(msg);
    /// msg.Close();
    /// Console.WriteLine(System.String.Concat("From:    [", message.From, "][", message.FromAddress, "]"));
    /// Console.WriteLine(System.String.Concat("To:      [", message.To, "]"));
    /// Console.WriteLine(System.String.Concat("Subject: [", message.Subject, "]"));
    /// Console.WriteLine(System.String.Concat("Date:    [", message.Date, "]"));
    /// Console.WriteLine(System.String.Concat("Body:    [", message.Body, "]"));
    /// </code>
    /// </example>
    public sealed class SharpMessage
    {
        private String _body = String.Empty;
        private DateTime _date;
        private String _from_addr = String.Empty;
        private String _from_name = String.Empty;
        private String _subject = String.Empty;

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpMessage" /> class based on the supplied <see cref="System.IO.Stream" />.
        /// </summary>
        /// <param name="message"><see cref="System.IO.Stream" /> that contains the message content</param>
        /// <remarks>The message content is automatically parsed.</remarks>
        public SharpMessage(Stream message)
            : this(message, true, true, null, null)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpMessage" /> class based on the supplied <see cref="System.IO.Stream" />.
        /// </summary>
        /// <param name="message"><see cref="System.IO.Stream" /> that contains the message content</param>
        /// <param name="attachments"><b>true</b> to allow attachments; <b>false</b> to skip them.</param>
        /// <param name="html"><b>true</b> to allow HTML content; <b>false</b> to ignore the html content.</param>
        /// <remarks>When the <b>attachments</b> parameter is true it's equivalent to adding <b>anmar.SharpMimeTools.MimeTopLevelMediaType.application</b>, <b>anmar.SharpMimeTools.MimeTopLevelMediaType.audio</b>, <b>anmar.SharpMimeTools.MimeTopLevelMediaType.image</b>, <b>anmar.SharpMimeTools.MimeTopLevelMediaType.video</b> to the allowed <see cref="anmar.SharpMimeTools.MimeTopLevelMediaType" />.<br />
        /// <b>anmar.SharpMimeTools.MimeTopLevelMediaType.text</b>, <b>anmar.SharpMimeTools.MimeTopLevelMediaType.multipart</b> and <b>anmar.SharpMimeTools.MimeTopLevelMediaType.message</b> are allowed in any case.<br /><br />
        /// In order to have better control over what is parsed, see the other contructors.
        /// </remarks>
        public SharpMessage(Stream message, bool attachments, bool html)
            : this(message, attachments, html, null, null)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpMessage" /> class based on the supplied <see cref="System.IO.Stream" />.
        /// </summary>
        /// <param name="message"><see cref="System.IO.Stream" /> that contains the message content</param>
        /// <param name="attachments"><b>true</b> to allow attachments; <b>false</b> to skip them.</param>
        /// <param name="html"><b>true</b> to allow HTML content; <b>false</b> to ignore the html content.</param>
        /// <param name="path">A <see cref="System.String" /> specifying the path on which to save the attachments found.</param>
        /// <remarks>When the <b>attachments</b> parameter is true it's equivalent to adding <b>anmar.SharpMimeTools.MimeTopLevelMediaType.application</b>, <b>anmar.SharpMimeTools.MimeTopLevelMediaType.audio</b>, <b>anmar.SharpMimeTools.MimeTopLevelMediaType.image</b>, <b>anmar.SharpMimeTools.MimeTopLevelMediaType.video</b> to the allowed <see cref="anmar.SharpMimeTools.MimeTopLevelMediaType" />.<br />
        /// <b>anmar.SharpMimeTools.MimeTopLevelMediaType.text</b>, <b>anmar.SharpMimeTools.MimeTopLevelMediaType.multipart</b> and <b>anmar.SharpMimeTools.MimeTopLevelMediaType.message</b> are allowed in any case.<br /><br />
        /// In order to have better control over what is parsed, see the other contructors.
        /// </remarks>
        public SharpMessage(Stream message, bool attachments, bool html, String path)
            : this(message, attachments, html, path, null)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpMessage" /> class based on the supplied <see cref="System.IO.Stream" />.
        /// </summary>
        /// <param name="message"><see cref="System.IO.Stream" /> that contains the message content</param>
        /// <param name="attachments"><b>true</b> to allow attachments; <b>false</b> to skip them.</param>
        /// <param name="html"><b>true</b> to allow HTML content; <b>false</b> to ignore the html content.</param>
        /// <param name="path">A <see cref="System.String" /> specifying the path on which to save the attachments found.</param>
        /// <param name="preferredtextsubtype">A <see cref="System.String" /> specifying the subtype to select for text parts that contain aternative content (plain, html, ...). Specify the <b>null</b> reference to maintain the default behavior (prefer whatever the originator sent as the preferred part). If there is not a text part with this subtype, the default one is used.</param>
        /// <remarks>When the <b>attachments</b> parameter is true it's equivalent to adding <b>anmar.SharpMimeTools.MimeTopLevelMediaType.application</b>, <b>anmar.SharpMimeTools.MimeTopLevelMediaType.audio</b>, <b>anmar.SharpMimeTools.MimeTopLevelMediaType.image</b>, <b>anmar.SharpMimeTools.MimeTopLevelMediaType.video</b> to the allowed <see cref="anmar.SharpMimeTools.MimeTopLevelMediaType" />.<br />
        /// <b>anmar.SharpMimeTools.MimeTopLevelMediaType.text</b>, <b>anmar.SharpMimeTools.MimeTopLevelMediaType.multipart</b> and <b>anmar.SharpMimeTools.MimeTopLevelMediaType.message</b> are allowed in any case.<br /><br />
        /// In order to have better control over what is parsed, see the other contructors.
        /// </remarks>
        public SharpMessage(Stream message, bool attachments, bool html, String path, String preferredtextsubtype)
        {
            MimeTopLevelMediaType types = MimeTopLevelMediaType.text | MimeTopLevelMediaType.multipart | MimeTopLevelMediaType.message;
            SharpDecodeOptions options = SharpDecodeOptions.None;
            if (attachments)
            {
                types = types | MimeTopLevelMediaType.application | MimeTopLevelMediaType.audio | MimeTopLevelMediaType.image | MimeTopLevelMediaType.video;
                options = options | SharpDecodeOptions.AllowAttachments;
            }
            if (html)
                options = options | SharpDecodeOptions.AllowHtml;
            if (path == null || !Directory.Exists(path))
                ParseMessage(message, types, options, preferredtextsubtype, null);
            else
                ParseMessage(message, types, options, preferredtextsubtype, path);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpMessage" /> class based on the supplied <see cref="System.IO.Stream" />.
        /// </summary>
        /// <param name="message"><see cref="System.IO.Stream" /> that contains the message content</param>
        /// <param name="types">A <see cref="anmar.SharpMimeTools.MimeTopLevelMediaType" /> value that specifies the allowed Mime-Types to being decoded.</param>
        /// <param name="html"><b>true</b> to allow HTML content; <b>false</b> to ignore the html content.</param>
        public SharpMessage(Stream message, MimeTopLevelMediaType types, bool html)
            : this(message, types, html, null, null)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpMessage" /> class based on the supplied <see cref="System.IO.Stream" />.
        /// </summary>
        /// <param name="message"><see cref="System.IO.Stream" /> that contains the message content</param>
        /// <param name="types">A <see cref="anmar.SharpMimeTools.MimeTopLevelMediaType" /> value that specifies the allowed Mime-Types to being decoded.</param>
        /// <param name="html"><b>true</b> to allow HTML content; <b>false</b> to ignore the html content.</param>
        /// <param name="path">A <see cref="System.String" /> specifying the path on which to save the attachments found.</param>
        public SharpMessage(Stream message, MimeTopLevelMediaType types, bool html, String path)
            : this(message, types, html, path, null)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpMessage" /> class based on the supplied <see cref="System.IO.Stream" />.
        /// </summary>
        /// <param name="message"><see cref="System.IO.Stream" /> that contains the message content</param>
        /// <param name="types">A <see cref="anmar.SharpMimeTools.MimeTopLevelMediaType" /> value that specifies the allowed Mime-Types to being decoded.</param>
        /// <param name="html"><b>true</b> to allow HTML content; <b>false</b> to ignore the html content.</param>
        /// <param name="path">A <see cref="System.String" /> specifying the path on which to save the attachments found.</param>
        /// <param name="preferredtextsubtype">A <see cref="System.String" /> specifying the subtype to select for text parts that contain aternative content (plain, html, ...). Specify the <b>null</b> reference to maintain the default behavior (prefer whatever the originator sent as the preferred part). If there is not a text part with this subtype, the default one is used.</param>
        public SharpMessage(Stream message, MimeTopLevelMediaType types, bool html, String path, String preferredtextsubtype)
            : this(message, types, ((html) ? SharpDecodeOptions.Default : SharpDecodeOptions.AllowAttachments), path, preferredtextsubtype)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpMessage" /> class based on the supplied <see cref="System.IO.Stream" />.
        /// </summary>
        /// <param name="message"><see cref="System.IO.Stream" /> that contains the message content</param>
        /// <param name="options"><see cref="anmar.SharpMimeTools.SharpDecodeOptions" /> to determine how to do the decoding (must be combined as a bit map).</param>
        public SharpMessage(Stream message, SharpDecodeOptions options)
            : this(message, options, null, null)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpMessage" /> class based on the supplied <see cref="System.IO.Stream" />.
        /// </summary>
        /// <param name="message"><see cref="System.IO.Stream" /> that contains the message content</param>
        /// <param name="options"><see cref="anmar.SharpMimeTools.SharpDecodeOptions" /> to determine how to do the decoding (must be combined as a bit map).</param>
        /// <param name="path">A <see cref="System.String" /> specifying the path on which to save the attachments found.</param>
        public SharpMessage(Stream message, SharpDecodeOptions options, String path)
            : this(message, options, path, null)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpMessage" /> class based on the supplied <see cref="System.IO.Stream" />.
        /// </summary>
        /// <param name="message"><see cref="System.IO.Stream" /> that contains the message content</param>
        /// <param name="options"><see cref="anmar.SharpMimeTools.SharpDecodeOptions" /> to determine how to do the decoding (must be combined as a bit map).</param>
        /// <param name="path">A <see cref="System.String" /> specifying the path on which to save the attachments found.</param>
        /// <param name="preferredtextsubtype">A <see cref="System.String" /> specifying the subtype to select for text parts that contain aternative content (plain, html, ...). Specify the <b>null</b> reference to maintain the default behavior (prefer whatever the originator sent as the preferred part). If there is not a text part with this subtype, the default one is used.</param>
        public SharpMessage(Stream message, SharpDecodeOptions options, String path, String preferredtextsubtype)
        {
            MimeTopLevelMediaType types;
            if ((options & SharpDecodeOptions.AllowAttachments) == SharpDecodeOptions.AllowAttachments)
                types = MimeTopLevelMediaType.text | MimeTopLevelMediaType.multipart | MimeTopLevelMediaType.message | MimeTopLevelMediaType.application | MimeTopLevelMediaType.audio | MimeTopLevelMediaType.image | MimeTopLevelMediaType.video;
            else
                types = MimeTopLevelMediaType.text | MimeTopLevelMediaType.multipart | MimeTopLevelMediaType.message;
            if (path != null && (options & SharpDecodeOptions.CreateFolder) != SharpDecodeOptions.CreateFolder && !Directory.Exists(path))
            {
                path = null;
            }
            ParseMessage(message, types, options, preferredtextsubtype, path);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpMessage" /> class based on the supplied <see cref="System.IO.Stream" />.
        /// </summary>
        /// <param name="message"><see cref="System.IO.Stream" /> that contains the message content</param>
        /// <param name="types">A <see cref="anmar.SharpMimeTools.MimeTopLevelMediaType" /> value that specifies the allowed Mime-Types to being decoded.</param>
        /// <param name="options"><see cref="anmar.SharpMimeTools.SharpDecodeOptions" /> to determine how to do the decoding (must be combined as a bit map).</param>
        /// <param name="path">A <see cref="System.String" /> specifying the path on which to save the attachments found.</param>
        /// <param name="preferredtextsubtype">A <see cref="System.String" /> specifying the subtype to select for text parts that contain aternative content (plain, html, ...). Specify the <b>null</b> reference to maintain the default behavior (prefer whatever the originator sent as the preferred part). If there is not a text part with this subtype, the default one is used.</param>
        public SharpMessage(Stream message, MimeTopLevelMediaType types, SharpDecodeOptions options, String path, String preferredtextsubtype)
        {
            if (path != null && (options & SharpDecodeOptions.CreateFolder) != SharpDecodeOptions.CreateFolder && !Directory.Exists(path))
            {
                path = null;
            }
            ParseMessage(message, types, options, preferredtextsubtype, path);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpMessage" /> class based on the supplied <see cref="System.String" />.
        /// </summary>
        /// <param name="message"><see cref="System.String" /> with the message content</param>
        public SharpMessage(String message)
            : this(new MemoryStream(System.Text.Encoding.ASCII.GetBytes(message)))
        {
        }

        /// <summary>
        /// <see cref="System.Collections.ICollection" /> that contains the attachments found in this message.
        /// </summary>
        /// <remarks>Each attachment is a <see cref="SharpAttachment" /> instance.</remarks>
        public List<SharpAttachment> Attachments { get; private set; }

        /// <summary>
        /// Text body
        /// </summary>
        /// <remarks>If there are more than one text part in the message, they are concatenated.</remarks>
        public String Body
        {
            get { return _body; }
        }

        /// <summary>
        /// Collection of <see cref="anmar.SharpMimeTools.SharpMimeAddress" /> instances found in the <b>Cc</b> header field.
        /// </summary>
        public SharpMimeAddressCollection Cc
        {
            get { return SharpMimeAddressCollection.Parse(Headers.Cc); }
        }

        /// <summary>
        /// Date
        /// </summary>
        /// <remarks>If there is not a <b>Date</b> field present in the headers (or it has an invalid format) then
        /// the date is extrated from the last <b>Received</b> field. If neither of them are found,
        /// <b>System.Date.MinValue</b> is returned.</remarks>
        public DateTime Date
        {
            get { return _date; }
        }

        /// <summary>
        /// From's name
        /// </summary>
        public String From
        {
            get { return _from_name; }
        }

        /// <summary>
        /// From's e-mail address
        /// </summary>
        public String FromAddress
        {
            get { return _from_addr; }
        }

        /// <summary>
        /// Gets a value indicating whether the body contains html content
        /// </summary>
        /// <value><b>true</b> if the body contains html content; otherwise, <b>false</b></value>
        public bool HasHtmlBody { get; private set; }

        /// <summary>
        /// <see cref="anmar.SharpMimeTools.SharpMimeHeader" /> instance for this message that contains the raw content of the headers.
        /// </summary>
        public SharpMimeHeader Headers { get; private set; }

        /// <summary>
        /// <b>Message-ID</b> header
        /// </summary>
        public String MessageID
        {
            get { return Headers.MessageID; }
        }

        /// <summary>
        /// <b>Subject</b> field
        /// </summary>
        /// <remarks>The field body is automatically RFC 2047 decoded if it's necessary</remarks>
        public String Subject
        {
            get { return _subject; }
        }

        /// <summary>
        /// Collection of <see cref="anmar.SharpMimeTools.SharpMimeAddress" /> found in the <b>To</b> header field.
        /// </summary>
        public SharpMimeAddressCollection To { get; private set; }

        /// <summary>
        /// Returns the requested header field body.
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        /// <remarks>The value if present is uncommented and decoded (RFC 2047).<br />
        /// If the requested field is not present in this instance, <see cref="System.String.Empty" /> is returned instead.</remarks>
        public String GetHeaderField(String name)
        {
            if (Headers == null)
                return String.Empty;
            return Headers.GetHeaderField(name, String.Empty, true, true);
        }

        /// <summary>
        /// Set the URL used to reference embedded parts from the HTML body (as specified on RFC2392)
        /// </summary>
        /// <param name="attachmentsurl">URL used to reference embedded parts from the HTML body.</param>
        /// <remarks>The supplied URL will be replaced with the following tokens for each attachment:<br />
        /// <ul>
        ///  <li><b>[MessageID]</b>: Will be replaced with the <see cref="MessageID" /> of the current instance.</li>
        ///  <li><b>[ContentID]</b>: Will be replaced with the <see cref="anmar.SharpMimeTools.SharpAttachment.ContentID" /> of the attachment.</li>
        ///  <li><b>[Name]</b>: Will be replaced with the <see cref="anmar.SharpMimeTools.SharpAttachment.Name" /> of the attachment (or with <see cref="System.IO.FileInfo.Name" /> from <see cref="anmar.SharpMimeTools.SharpAttachment.SavedFile" /> if the instance has been already saved to disk).</li>
        /// </ul>
        ///</remarks>
        public void SetUrlBase(String attachmentsurl)
        {
            // Not a html boy or not body at all
            if (!HasHtmlBody || _body.Length == 0)
                return;
            // No references found, so nothing to do
            if (_body.IndexOf("cid:") == -1 && _body.IndexOf("mid:") == -1)
                return;
            String msgid = SharpMimeTools.Rfc2392Url(MessageID);
            // There is a base url and attachments, so try refererencing them properly
            if (attachmentsurl != null && Attachments != null && Attachments.Count > 0)
            {
                for (int i = 0, count = Attachments.Count; i < count; i++)
                {
                    SharpAttachment attachment = (SharpAttachment)Attachments[i];
                    if (attachment.ContentID != null)
                    {
                        String conid = SharpMimeTools.Rfc2392Url(attachment.ContentID);
                        if (conid.Length > 0)
                        {
                            if (_body.IndexOf("cid:" + conid) != -1)
                                _body = _body.Replace("cid:" + conid, ReplaceUrlTokens(attachmentsurl, attachment));
                            if (!String.IsNullOrEmpty(msgid) && _body.IndexOf(String.Format("mid:{0}/{1}", msgid, conid)) != -1)
                                _body = _body.Replace(String.Format("mid:{0}/{1}", msgid, conid), ReplaceUrlTokens(attachmentsurl, attachment));
                            // No more references found, so nothing to do
                            if (_body.IndexOf("cid:") == -1 && _body.IndexOf("mid:") == -1)
                                break;
                        }
                    }
                }
            }
            // The rest references must be to text parts
            // so rewrite them to refer to the named anchors added by ParseMessage
            if (_body.IndexOf("cid:") != -1)
            {
                _body = _body.Replace("cid:", String.Format("#{0}_", msgid));
            }
            if (msgid.Length > 0 && _body.IndexOf("mid:") != -1)
            {
                _body = _body.Replace(String.Format("mid:{0}/", msgid), String.Format("#{0}_", msgid));
                _body = _body.Replace("mid:" + msgid, "#" + msgid);
            }
        }

        private void ParseMessage(Stream stream, MimeTopLevelMediaType types, SharpDecodeOptions options, String preferredtextsubtype, String path)
        {
            Attachments = new List<SharpAttachment>();
            using (SharpMimeMessage message = new SharpMimeMessage(stream))
            {
                ParseMessage(message, types, (options & SharpDecodeOptions.AllowHtml) == SharpDecodeOptions.AllowHtml, options, preferredtextsubtype, path);
                Headers = message.Header;
            }
            // find and decode uuencoded content if configured to do so (and attachments a allowed)
            if ((options & SharpDecodeOptions.UuDecode) == SharpDecodeOptions.UuDecode
                   && (options & SharpDecodeOptions.AllowAttachments) == SharpDecodeOptions.AllowAttachments)
                UuDecode(path);
            // Date
            _date = SharpMimeTools.parseDate(Headers.Date);
            if (_date.Equals(DateTime.MinValue))
            {
                String date = Headers["Received"];
                if (date == null)
                    date = String.Empty;
                if (date.IndexOf("\r\n") > 0)
                    date = date.Substring(0, date.IndexOf("\r\n"));
                if (date.LastIndexOf(';') > 0)
                    date = date.Substring(date.LastIndexOf(';') + 1).Trim();
                else
                    date = String.Empty;
                _date = SharpMimeTools.parseDate(date);
            }
            // Subject
            _subject = SharpMimeTools.parserfc2047Header(Headers.Subject);
            // To
            To = SharpMimeAddressCollection.Parse(Headers.To);
            // From
            SharpMimeAddressCollection from = SharpMimeAddressCollection.Parse(Headers.From);
            foreach (SharpMimeAddress item in from)
            {
                _from_name = item["name"];
                _from_addr = item["address"];
                if (_from_name == null || _from_name.Equals(String.Empty))
                    _from_name = item["address"];
            }
        }

        private void ParseMessage(SharpMimeMessage part, MimeTopLevelMediaType types, bool html, SharpDecodeOptions options, String preferredtextsubtype, String path)
        {
            if ((types & part.Header.TopLevelMediaType) != part.Header.TopLevelMediaType)
            {
                return;
            }
            switch (part.Header.TopLevelMediaType)
            {
                case MimeTopLevelMediaType.multipart:
                case MimeTopLevelMediaType.message:
                    // TODO: allow other subtypes of "message"
                    if (part.Header.TopLevelMediaType.Equals(MimeTopLevelMediaType.message))
                    {
                        // Only message/rfc822 allowed, other subtypes ignored
                        if (part.Header.SubType == "rfc822")
                        {
                            // If NotRecursiveRfc822 option is set, handle part as an attachment
                            if ((options & SharpDecodeOptions.NotRecursiveRfc822) == SharpDecodeOptions.NotRecursiveRfc822)
                            {
                                goto case anmar.SharpMimeTools.MimeTopLevelMediaType.application;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                    if (part.Header.SubType.Equals("alternative"))
                    {
                        if (part.PartsCount > 0)
                        {
                            SharpMimeMessage altenative = null;
                            // Get the first mime part of the alternatives that has a accepted Mime-Type
                            for (int i = part.PartsCount; i > 0; i--)
                            {
                                SharpMimeMessage item = part.GetPart(i - 1);
                                if ((types & part.Header.TopLevelMediaType) != part.Header.TopLevelMediaType
                                    || (!html && item.Header.TopLevelMediaType.Equals(MimeTopLevelMediaType.text) && item.Header.SubType.Equals("html"))
                                   )
                                {
                                    continue;
                                }
                                // First allowed one.
                                if (altenative == null)
                                {
                                    altenative = item;
                                    // We don't have to select body part based on subtype if not asked for, or not a text one
                                    // or it's already the preferred one
                                    if (preferredtextsubtype == null || item.Header.TopLevelMediaType != MimeTopLevelMediaType.text || (preferredtextsubtype != null && item.Header.SubType == preferredtextsubtype))
                                    {
                                        break;
                                    }
                                    // This one is preferred over the last part
                                }
                                else if (preferredtextsubtype != null && item.Header.TopLevelMediaType == MimeTopLevelMediaType.text && item.Header.SubType == preferredtextsubtype)
                                {
                                    altenative = item;
                                    break;
                                }
                            }
                            if (altenative != null)
                            {
                                // If message body as html is allowed and part has a Content-ID field
                                // add an anchor to mark this body part
                                if (html && part.Header.Contains("Content-ID") && (options & SharpDecodeOptions.NamedAnchors) == SharpDecodeOptions.NamedAnchors)
                                {
                                    // There is a previous text body, so enclose it in <pre>
                                    if (!HasHtmlBody && _body.Length > 0)
                                    {
                                        _body = String.Concat("<pre>", HttpUtility.HtmlEncode(_body), "</pre>");
                                        HasHtmlBody = true;
                                    }
                                    // Add the anchor
                                    _body = String.Concat(_body, "<a name=\"", SharpMimeTools.Rfc2392Url(MessageID), "_", SharpMimeTools.Rfc2392Url(part.Header.ContentID), "\"></a>");
                                }
                                ParseMessage(altenative, types, html, options, preferredtextsubtype, path);
                            }
                        }
                        // TODO: Take into account each subtype of "multipart" and "message"
                    }
                    else if (part.PartsCount > 0)
                    {
                        foreach (SharpMimeMessage item in part)
                        {
                            ParseMessage(item, types, html, options, preferredtextsubtype, path);
                        }
                    }
                    break;
                case MimeTopLevelMediaType.text:
                    if ((part.Disposition == null || !part.Disposition.Equals("attachment"))
                        && (part.Header.SubType.Equals("plain") || part.Header.SubType.Equals("html")))
                    {
                        bool body_was_html = HasHtmlBody;
                        // HTML content not allowed
                        if (part.Header.SubType.Equals("html"))
                        {
                            if (!html)
                                break;
                            else
                                HasHtmlBody = true;
                        }
                        if (html && part.Header.Contains("Content-ID") && (options & SharpDecodeOptions.NamedAnchors) == SharpDecodeOptions.NamedAnchors)
                        {
                            HasHtmlBody = true;
                        }
                        if (HasHtmlBody && !body_was_html && !String.IsNullOrEmpty(_body))
                        {
                            _body = String.Concat("<pre>", HttpUtility.HtmlEncode(_body), "</pre>");
                        }
                        // If message body is html and this part has a Content-ID field
                        // add an anchor to mark this body part
                        if (HasHtmlBody && part.Header.Contains("Content-ID") && (options & SharpDecodeOptions.NamedAnchors) == SharpDecodeOptions.NamedAnchors)
                        {
                            _body = String.Concat(_body, "<a name=\"", SharpMimeTools.Rfc2392Url(MessageID), "_", SharpMimeTools.Rfc2392Url(part.Header.ContentID), "\"></a>");
                        }
                        if (HasHtmlBody && part.Header.SubType.Equals("plain"))
                        {
                            _body = String.Concat(_body, "<pre>", HttpUtility.HtmlEncode(part.BodyDecoded), "</pre>");
                        }
                        else
                            _body = String.Concat(_body, part.BodyDecoded);
                    }
                    else
                    {
                        if ((types & MimeTopLevelMediaType.application) != MimeTopLevelMediaType.application)
                        {
                            return;
                        }
                        goto case anmar.SharpMimeTools.MimeTopLevelMediaType.application;
                    }
                    break;
                case MimeTopLevelMediaType.application:
                case MimeTopLevelMediaType.audio:
                case MimeTopLevelMediaType.image:
                case MimeTopLevelMediaType.video:
                    // Attachments not allowed.
                    if ((options & SharpDecodeOptions.AllowAttachments) != SharpDecodeOptions.AllowAttachments)
                        break;
                    SharpAttachment attachment = null;
                    // Save to a file
                    if (path != null)
                    {
                        FileInfo file = part.DumpBody(path, true);
                        if (file != null)
                        {
                            attachment = new SharpAttachment(file);
                            attachment.Name = file.Name;
                            attachment.Size = file.Length;
                        }
                        // Save to a stream
                    }
                    else
                    {
                        using (MemoryStream stream = new MemoryStream())
                        {
                            if (part.DumpBody(stream))
                            {
                                if (stream != null && stream.CanSeek)
                                    stream.Seek(0, SeekOrigin.Begin);
                                attachment = new SharpAttachment(stream);
                                if (part.Name != null)
                                    attachment.Name = part.Name;
                                else
                                    attachment.Name = String.Concat("generated_", part.GetHashCode(), ".", part.Header.SubType);
                                attachment.Size = stream.Length;
                            }
                        }
                    }
                    if (attachment != null && part.Header.SubType == "ms-tnef" && (options & SharpDecodeOptions.DecodeTnef) == SharpDecodeOptions.DecodeTnef)
                    {
                        // Try getting attachments form a tnef stream
                        Stream stream = attachment.Stream;
                        SharpTnefMessage tnef = new SharpTnefMessage(stream);
                        if (tnef.Parse(path))
                        {
                            if (tnef.Attachments != null)
                            {
                                Attachments.AddRange(tnef.Attachments);
                            }
                            attachment.Close();
                            // Delete the raw tnef file
                            if (attachment.SavedFile != null)
                            {
                                if (stream != null && stream.CanRead)
                                {
                                    stream.Close();
                                    stream = null;
                                }
                                attachment.SavedFile.Delete();
                            }
                            attachment = null;
                            tnef.Close();
                        }
                        else
                        {
                            // The read-only stream is no longer needed and locks the file
                            if (attachment.SavedFile != null && stream != null && stream.CanRead)
                            {
                                stream.Close();
                                stream = null;
                            }
                        }
                        stream = null;
                        tnef = null;
                    }
                    if (attachment != null)
                    {
                        if (part.Disposition != null && part.Disposition == "inline")
                        {
                            attachment.Inline = true;
                        }
                        attachment.MimeTopLevelMediaType = part.Header.TopLevelMediaType;
                        attachment.MimeMediaSubType = part.Header.SubType;
                        // Store attachment's CreationTime
                        if (part.Header.ContentDispositionParameters.ContainsKey("creation-date"))
                            attachment.CreationTime = SharpMimeTools.parseDate(part.Header.ContentDispositionParameters["creation-date"]);
                        // Store attachment's LastWriteTime
                        if (part.Header.ContentDispositionParameters.ContainsKey("modification-date"))
                            attachment.LastWriteTime = SharpMimeTools.parseDate(part.Header.ContentDispositionParameters["modification-date"]);
                        if (part.Header.Contains("Content-ID"))
                            attachment.ContentID = part.Header.ContentID;
                        Attachments.Add(attachment);
                    }
                    break;
                default:
                    break;
            }
        }

        private String ReplaceUrlTokens(String url, SharpAttachment attachment)
        {
            if (url == null || url.Length == 0 || url.IndexOf('[') == -1 || url.IndexOf(']') == -1)
                return url;
            if (url.IndexOf("[MessageID]") != -1)
            {
                url = url.Replace("[MessageID]", HttpUtility.UrlEncode(SharpMimeTools.Rfc2392Url(MessageID)));
            }
            if (attachment != null && attachment.ContentID != null)
            {
                if (url.IndexOf("[ContentID]") != -1)
                {
                    url = url.Replace("[ContentID]", HttpUtility.UrlEncode(SharpMimeTools.Rfc2392Url(attachment.ContentID)));
                }
                if (url.IndexOf("[Name]") != -1)
                {
                    if (attachment.SavedFile != null)
                    {
                        url = url.Replace("[Name]", HttpUtility.UrlEncode(attachment.SavedFile.Name));
                    }
                    else
                    {
                        url = url.Replace("[Name]", HttpUtility.UrlEncode(attachment.Name));
                    }
                }
            }
            return url;
        }

        private void UuDecode(String path)
        {
            if (_body.Length == 0 || _body.IndexOf("begin ") == -1 || _body.IndexOf("end") == -1)
                return;
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            using (StringReader reader = new StringReader(_body))
            {
                Stream stream = null;
                SharpAttachment attachment = null;
                for (String line = reader.ReadLine(); line != null; line = reader.ReadLine())
                {
                    if (stream == null)
                    {
                        // Found the start point of uuencoded content
                        if (line.Length > 10 && line[0] == 'b' && line[1] == 'e' && line[2] == 'g' && line[3] == 'i' && line[4] == 'n' && line[5] == ' ' && line[9] == ' ')
                        {
                            String name = Path.GetFileName(line.Substring(10));
                            // In-Memory decoding
                            if (path == null)
                            {
                                attachment = new SharpAttachment();
                                stream = attachment.Stream;
                                // Filesystem decoding
                            }
                            else
                            {
                                attachment = new SharpAttachment(new FileInfo(Path.Combine(path, name)));
                                stream = attachment.SavedFile.OpenWrite();
                            }
                            attachment.Name = name;
                            // Not uuencoded line, so add it to new body
                        }
                        else
                        {
                            sb.Append(line);
                            sb.Append(Environment.NewLine);
                        }
                    }
                    else
                    {
                        // Content finished
                        if (line.Length == 3 && line == "end")
                        {
                            stream.Flush();
                            if (stream.Length > 0)
                            {
                                attachment.Size = stream.Length;
                                Attachments.Add(attachment);
                            }
                            // When decoding to a file, close the stream.
                            if (attachment.SavedFile != null || stream.Length == 0)
                            {
                                stream.Close();
                            }
                            attachment = null;
                            stream = null;
                            // Decode and write uuencoded line
                        }
                        else
                        {
                            SharpMimeTools.UuDecodeLine(line, stream);
                        }
                    }
                }
                if (stream != null && stream.CanRead)
                {
                    stream.Close();
                    stream = null;
                }
            }
            _body = sb.ToString();
            sb = null;
        }
    }
}
