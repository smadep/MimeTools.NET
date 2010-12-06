using System;
using System.Collections;

namespace anmar.SharpMimeTools
{
    internal class SharpMimeAddressCollection : IEnumerable
    {
        protected ArrayList list = new ArrayList();

        public SharpMimeAddressCollection(String text)
        {
            String[] tokens = ABNF.address_regex.Split(text);
            foreach (String token in tokens)
            {
                if (ABNF.address_regex.IsMatch(token))
                    Add(new SharpMimeAddress(token));
            }
        }
        
        public SharpMimeAddress this[int index]
        {
            get
            {
                return Get(index);
            }
        }
        
        public IEnumerator GetEnumerator()
        {
            return list.GetEnumerator();
        }
        
        public void Add(SharpMimeAddress address)
        {
            list.Add(address);
        }
        
        public SharpMimeAddress Get(int index)
        {
            return (SharpMimeAddress)list[index];
        }
        
        public static SharpMimeAddressCollection Parse(String text)
        {
            if (text == null)
                throw new ArgumentNullException();
            return new SharpMimeAddressCollection(text);
        }
        
        public int Count
        {
            get
            {
                return list.Count;
            }
        }
        
        public override string ToString()
        {
            System.Text.StringBuilder text = new System.Text.StringBuilder();
            foreach (SharpMimeAddress token in list)
            {
                text.Append(token.ToString());
                if (token.Length > 0)
                    text.Append("; ");
            }
            return text.ToString();
        }
    }
}
