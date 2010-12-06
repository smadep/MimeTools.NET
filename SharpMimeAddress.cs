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

namespace anmar.SharpMimeTools
{
    /// <summary>
    /// rfc 2822 email address
    /// </summary>
    public class SharpMimeAddress
    {
        private readonly String name;
        private readonly String address;

        /// <summary>
        /// Initializes a new address from a RFC 2822 name-addr specification string
        /// </summary>
        /// <param name="dir">RFC 2822 name-addr address</param>
        /// 
        public SharpMimeAddress(String dir)
        {
            name = SharpMimeTools.parseFrom(dir, 1);
            address = SharpMimeTools.parseFrom(dir, 2);
        }
        
        /// <summary>
        /// Gets the decoded address or name contained in the name-addr
        /// </summary>
        public String this[Object key]
        {
            get
            {
                if (key == null)
                    throw new ArgumentNullException();
                switch (key.ToString())
                {
                    case "0":
                    case "name":
                        return name;
                    case "1":
                    case "address":
                        return address;
                }
                return null;
            }
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override String ToString()
        {
            if (name.Equals(String.Empty) && address.Equals(String.Empty))
                return "";
            if (name.Equals(String.Empty))
                return String.Format("<{0}>", address);
            else
                return String.Format("\"{0}\" <{1}>", name, address);
        }
        
        /// <summary>
        /// Gets the length of the decoded address
        /// </summary>
        public int Length
        {
            get
            {
                return name.Length + address.Length;
            }
        }
    }
}
