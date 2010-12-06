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
using System.IO;
using System.Text;

namespace anmar.SharpMimeTools
{
    /// <summary>
    /// </summary>
    internal class SharpMimeMessageStream
    {
#if LOG
		private static log4net.ILog log  = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
#endif
        protected Stream stream;
        private StreamReader sr;
        private Encoding enc;
        protected long initpos;
        protected long finalpos;

        private String _buf;
        private long _buf_initpos;
        private long _buf_finalpos;

        public SharpMimeMessageStream(Stream stream)
        {
            this.stream = stream;
            enc = SharpMimeHeader.EncodingDefault;
            sr = new StreamReader(this.stream, enc);
        }
        
        public SharpMimeMessageStream(Byte[] buffer)
        {
            stream = new MemoryStream(buffer);
            enc = SharpMimeHeader.EncodingDefault;
            sr = new StreamReader(stream, enc);
        }
        
        public void Close()
        {
            sr.Close();
        }
        
        public String ReadAll()
        {
            return ReadLines(Position, stream.Length);
        }
        
        public String ReadAll(long start)
        {
            return ReadLines(start, stream.Length);
        }

        public String ReadLine()
        {
            String line = null;
            if (_buf != null)
            {
                line = _buf;
                initpos = _buf_initpos;
                finalpos = _buf_finalpos;
                _buf = null;
            }
            else
            {
                StringBuilder sb = new StringBuilder(80);
                int ending = 0;
                initpos = Position;
                for (int current = sr.Read(); current != -1; current = sr.Read())
                {
                    sb.Append((char)current);
                    if (current == '\r')
                        ending++;
                    else if (current == '\n')
                    {
                        ending++;
                        break;
                    }
                }
                // Line ending found
                if (ending > 0)
                {
                    // Bytes read
                    finalpos += enc.GetByteCount(sb.ToString());
                    // A single dot is treated as message end
                    if (sb.Length == (1 + ending) && sb[0] == '.')
                        sb = null;
                    // Undo the double dots
                    else if (sb.Length > (1 + ending) && sb[0] == '.' && sb[1] == '.')
                        sb.Remove(0, 1);
                    if (sb != null)
                        line = sb.ToString(0, sb.Length - ending);
                }
                else
                {
                    // Not line ending found, so we are at the end of the stream
                    finalpos = stream.Length;
                    // though at the end of the stream there may be some content
                    if (sb.Length > 0)
                        line = sb.ToString();
                }
                sb = null;
            }
            return line;
        }
        
        public String ReadLines(long start, long end)
        {
            return ReadLinesSB(start, end).ToString();
        }

        public StringBuilder ReadLinesSB(long start, long end)
        {
            StringBuilder lines = new StringBuilder();
            String line;
            SeekPoint(start);
            do
            {
                line = ReadLine();
                if (line != null)
                {
                    // TODO: try catch
                    if (lines.Length > 0)
                        lines.Append(ABNF.CRLF);
                    lines.Append(line);
                }
            } while (line != null && Position != -1 && Position < end);
            initpos = start;
            return lines;
        }
        
        public void ReadLine_Undo()
        {
            SeekPoint(initpos);
            finalpos = initpos;
        }
        
        public void ReadLine_Undo(String line)
        {
            _buf_initpos = initpos;
            _buf_finalpos = finalpos;
            _buf = line;
            finalpos = initpos;
        }

        public String ReadUnfoldedLine()
        {
            long initpos = Position;
            String first_line = ReadLine();
            if (first_line != null && first_line.Length > 0)
            {
                StringBuilder line = null;
                String tmpline;
                for (; ; )
                {
                    tmpline = ReadLine();
                    // RFC 2822 - 2.2.3 Long Header Fields
                    if (tmpline != null && tmpline.Length > 0 && (tmpline[0] == ' ' || tmpline[0] == '\t'))
                    {
                        if (line == null)
                            line = new StringBuilder(first_line, 72);
                        line.Append(tmpline);
                    }
                    else
                    {
                        ReadLine_Undo(tmpline);
                        break;
                    }
                }
                this.initpos = initpos;
                if (finalpos != this.initpos)
                {
                    if (line == null)
                        return first_line;
                    else
                        return line.ToString();
                }
                else
                    return null;
            }
            return (finalpos != this.initpos) ? first_line : null;
        }
        
        public void SaveTo(Stream stream, long start, long end)
        {
            if (start < 0 || stream == null || !stream.CanWrite)
                return;
            SeekPoint(start);
            if (end == -1)
            {
                end = this.stream.Length;
            }
            int n = 0;
            long pending = end - start;
            if (pending <= 0)
                return;
            byte[] buffer = new byte[(pending > 4 * 1024) ? 4 * 1024 : pending];
            do
            {
                n = this.stream.Read(buffer, 0, (pending > buffer.Length) ? buffer.Length : (int)pending);
                if (n > 0)
                {
                    pending -= n;
                    if (pending == 0)
                    {
                        if (buffer[n - 1] == '\n')
                        {
                            n--;
                        }
                        if (n > 0 && buffer[n - 1] == '\r')
                        {
                            n--;
                        }
                    }
                    if (n > 0)
                        stream.Write(buffer, 0, n);

                }
            } while (n > 0);
        }
        
        public bool SeekLine(long line)
        {
            long linenumber = 0;
            SeekOrigin();
            for (; linenumber < (line - 1) && ReadLine() != null; linenumber++) { }
            return (linenumber == (line - 1)) ? true : false;
        }
        
        public void SeekOrigin()
        {
            SeekPoint(0);
        }
        
        public void SeekPoint(long point)
        {
            if (sr.BaseStream.CanSeek && sr.BaseStream.Seek(point, System.IO.SeekOrigin.Begin) != point)
            {
#if LOG
				if ( log.IsErrorEnabled) log.Error ("Error while seeking");
#endif
                throw new IOException();
            }
            else
            {
                sr.DiscardBufferedData();
                finalpos = point;
            }
            _buf = null;
        }

        public Encoding Encoding
        {
            set
            {
                if (value != null && enc.CodePage != value.CodePage)
                {
                    enc = value;
                    SeekPoint(Position);
                    sr = new StreamReader(stream, enc);
                }
            }
        }
        
        public long Position
        {
            get { return finalpos; }
        }
        
        public long Position_preRead
        {
            get { return initpos; }
        }
        
        public Stream Stream
        {
            get { return stream; }
        }
    }
}
