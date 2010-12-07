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

namespace anmar.SharpMimeTools
{
    internal class SharpMimeMessageCollection : IEnumerable
    {
        protected ArrayList messages = new ArrayList();

        public SharpMimeMessage this[int index]
        {
            get { return Get(index); }
        }

        public int Count
        {
            get
            {
                return messages.Count;
            }
        }

        public SharpMimeMessage Parent { get; set; }

        public void Add(SharpMimeMessage msg)
        {
            messages.Add(msg);
        }

        public SharpMimeMessage Get(int index)
        {
            return (SharpMimeMessage)messages[index];
        }

        public IEnumerator GetEnumerator()
        {
            return messages.GetEnumerator();
        }
        
        public void Clear()
        {
            messages.Clear();
        }
    }
}
