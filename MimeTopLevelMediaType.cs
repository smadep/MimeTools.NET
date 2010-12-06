using System;

namespace anmar.SharpMimeTools
{
    /// <summary>
    /// RFC 2046 Initial top-level media types
    /// </summary>
    [Flags]
    public enum MimeTopLevelMediaType
    {
        /// <summary>
        /// RFC 2046 section 4.1
        /// </summary>
        text = 1,
        /// <summary>
        /// RFC 2046 section 4.2
        /// </summary>
        image = 2,
        /// <summary>
        /// RFC 2046 section 4.3
        /// </summary>
        audio = 4,
        /// <summary>
        /// RFC 2046 section 4.4
        /// </summary>
        video = 8,
        /// <summary>
        /// RFC 2046 section 4.5
        /// </summary>
        application = 16,
        /// <summary>
        /// RFC 2046 section 5.1
        /// </summary>
        multipart = 32,
        /// <summary>
        /// RFC 2046 section 5.2
        /// </summary>
        message = 64
    }
}
