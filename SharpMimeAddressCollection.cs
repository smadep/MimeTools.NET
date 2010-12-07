using System;
using System.Collections;
using System.Collections.Generic;

namespace anmar.SharpMimeTools
{
    public class SharpMimeAddressCollection : List<SharpMimeAddress>
    {
        public SharpMimeAddressCollection(String text)
        {
            String[] tokens = ABNF.address_regex.Split(text);
            foreach (String token in tokens)
            {
                if (ABNF.address_regex.IsMatch(token))
                    Add(new SharpMimeAddress(token));
            }
        }

        public SharpMimeAddressCollection()
        {
            
        }

        public SharpMimeAddressCollection(int capacity)
            : base(capacity)
        {
            
        }

        public SharpMimeAddressCollection(IEnumerable<SharpMimeAddress> collection)
            : base(collection)
        {
            
        }

        public static SharpMimeAddressCollection Parse(String text)
        {
            if (text == null)
                throw new ArgumentNullException();
            return new SharpMimeAddressCollection(text);
        }
        
        public override string ToString()
        {
            System.Text.StringBuilder text = new System.Text.StringBuilder();
            foreach (SharpMimeAddress token in this)
            {
                text.Append(token.ToString());
                if (token.Length > 0)
                    text.Append("; ");
            }
            return text.ToString();
        }
    }
}
