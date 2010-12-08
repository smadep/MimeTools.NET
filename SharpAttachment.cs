using System;
using System.IO;

namespace anmar.SharpMimeTools
{
    /// <summary>
    /// This class provides the basic functionality for handling attachments
    /// </summary>
    public class SharpAttachment
    {
        private DateTime _ctime = DateTime.MinValue;
        private DateTime _mtime = DateTime.MinValue;
        private String _name;
        private MemoryStream _stream;

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpAttachment" /> class based on the supplied <see cref="System.IO.MemoryStream" />.
        /// </summary>
        /// <param name="stream"><see cref="System.IO.MemoryStream" /> instance that contains the attachment content.</param>
        public SharpAttachment(MemoryStream stream)
        {
            _stream = stream;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpAttachment" /> class.
        /// </summary>
        public SharpAttachment()
        {
            _stream = new MemoryStream();
        }
        
        /// <summary>
        /// Initializes a new instance of the <see cref="anmar.SharpMimeTools.SharpAttachment" /> class based on the supplied <see cref="System.IO.FileInfo" />.
        /// </summary>
        /// <param name="file"><see cref="System.IO.MemoryStream" /> instance that contains the info about the attachment.</param>
        public SharpAttachment(FileInfo file)
        {
            SavedFile = file;
            _ctime = file.CreationTime;
            _mtime = file.LastWriteTime;
        }

        /// <summary>
        /// Gets or sets the Content-ID of this attachment.
        /// </summary>
        /// <value>Content-ID header of this instance. Or the <b>null</b> reference.</value>
        public String ContentID { get; set; }

        /// <summary>
        /// Gets or sets the time when the file associated with this attachment was created.
        /// </summary>
        /// <value>The time this attachment was last written to.</value>
        public DateTime CreationTime
        {
            get { return _ctime; }
            set { _ctime = value; }
        }

        /// <summary>
        /// Gets or sets value indicating whether the current instance is an inline attachment.
        /// </summary>
        /// <value><b>true</b> is it's an inline attachment; <b>false</b> otherwise.</value>
        public bool Inline { get; set; }

        /// <summary>
        /// Gets or sets the time when the file associated with this attachment was last written to.
        /// </summary>
        /// <value>The time this attachment was last written to.</value>
        public DateTime LastWriteTime
        {
            get { return _mtime; }
            set { _mtime = value; }
        }

        /// <summary>
        /// Gets or sets the name of the attachment.
        /// </summary>
        /// <value>The name of the attachment.</value>
        public String Name
        {
            get { return _name; }
            set
            {
                String name = SharpMimeTools.GetFileName(value);
                if (value != null && name == null && _name != null && Path.HasExtension(value))
                {
                    name = Path.ChangeExtension(_name, Path.GetExtension(value));
                }
                _name = name;
            }
        }

        /// <summary>
        /// Gets or sets top-level media type for this <see cref="SharpAttachment" /> instance
        /// </summary>
        /// <value>Top-level media type from Content-Type header field of this <see cref="SharpAttachment" /> instance</value>
        public String MimeMediaSubType { get; set; }

        /// <summary>
        /// Gets or sets SubType for this <see cref="SharpAttachment" /> instance
        /// </summary>
        /// <value>SubType from Content-Type header field of this <see cref="SharpAttachment" /> instance</value>
        public MimeTopLevelMediaType MimeTopLevelMediaType { get; set; }

        /// <summary>
        /// Gets or sets size (in bytes) for this <see cref="SharpAttachment" /> instance.
        /// </summary>
        /// <value>Size of this <see cref="SharpAttachment" /> instance</value>
        public long Size { get; set; }

        /// <summary>
        /// Gets the <see cref="System.IO.FileInfo" /> of the saved file.
        /// </summary>
        /// <value>The <see cref="System.IO.FileInfo" /> of the saved file.</value>
        public FileInfo SavedFile { get; private set; }

        /// <summary>
        /// Gets the <see cref="System.IO.Stream " /> of the attachment.
        /// </summary>
        /// <value>The <see cref="System.IO.Stream " /> of the attachment.</value>
        /// <remarks>If the underling stream exists, it's returned. If the file has been saved, it opens <see cref="SavedFile" /> for reading.</remarks>
        public Stream Stream
        {
            get
            {
                if (_stream != null)
                    return _stream;
                else if (SavedFile != null)
                    return SavedFile.OpenRead();
                else
                    return null;
            }
        }
        
        /// <summary>
        /// Closes the underling stream if it's open.
        /// </summary>
        public void Close()
        {
            if (_stream != null && _stream.CanRead)
                _stream.Close();
            _stream = null;
        }
        
        /// <summary>
        /// Saves of the attachment to a file in the given path.
        /// </summary>
        /// <param name="path">A <see cref="System.String" /> specifying the path on which to save the attachment.</param>
        /// <param name="overwrite"><b>true</b> if the destination file can be overwritten; otherwise, <b>false</b>.</param>
        /// <returns><see cref="System.IO.FileInfo" /> of the saved file. <b>null</b> when it fails to save.</returns>
        /// <remarks>If the file was already saved, the previous <see cref="System.IO.FileInfo" /> is returned.<br />
        /// Once the file is successfully saved, the stream is closed and <see cref="SavedFile" /> property is updated.</remarks>
        public FileInfo Save(String path, bool overwrite)
        {
            if (path == null || _name == null)
                return null;
            if (_stream == null)
            {
                if (SavedFile != null)
                    return SavedFile;
                else
                    return null;
            }
            if (!_stream.CanRead)
            {
                return null;
            }
            FileInfo file = new FileInfo(Path.Combine(path, _name));
            if (!file.Directory.Exists)
            {
                return null;
            }
            if (file.Exists)
            {
                if (overwrite)
                {
                    try
                    {
                        file.Delete();
                    }
                    catch (Exception)
                    {
                        return null;
                    }
                }
                else
                {
                    // Though the file already exists, we set the times
                    if (_mtime != DateTime.MinValue && file.LastWriteTime != _mtime)
                        file.LastWriteTime = _mtime;
                    if (_ctime != DateTime.MinValue && file.CreationTime != _ctime)
                        file.CreationTime = _ctime;
                    return null;
                }
            }
            try
            {
                FileStream stream = file.OpenWrite();
                _stream.WriteTo(stream);
                stream.Flush();
                stream.Close();
                Close();
                if (_mtime != DateTime.MinValue)
                    file.LastWriteTime = _mtime;
                if (_ctime != DateTime.MinValue)
                    file.CreationTime = _ctime;
                SavedFile = file;
            }
            catch (Exception)
            {
                return null;
            }
            return file;
        }
    }
}
