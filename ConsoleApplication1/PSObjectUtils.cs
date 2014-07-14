using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;

namespace Exchange
{
    class PSObjectUtils
    {
        public bool GetBoolean(PSObject obj, string name)
        {
            if (!IsHashRefNull(obj, name))
            {
                return (bool)obj.Properties[name].Value;
            }
            return false;
        }

        public string GetString(PSObject obj, string name)
        {
            if (!IsHashRefNull(obj, name))
            {
                Object value = obj.Properties[name].Value;
                if (value is PSObject)
                {
                    PSObject innerObj = (PSObject)value;
                    return innerObj.ToString();
                }
                else if (value is string)
                {
                    return (string)value;
                }
                else if (value is ArrayList)
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (Object innerObj in (ArrayList)value)
                    {
                        sb.Append(innerObj.ToString());
                    }
                    return sb.ToString();
                }
                else
                {
                    return value.ToString();
                }
            }
            return "";
        }

        public Object[] GetObjectArray(PSObject obj, string name)
        {
            if (!IsHashRefNull(obj, name))
            {
                return (Object[])obj.Properties[name].Value;
            }
            return new Object[] { };
        }

        public HashSet<string> GetHashSet(PSObject obj, string name)
        {
            if (!IsHashRefNull(obj, name))
            {
                return (HashSet<string>)obj.Properties[name].Value;
            }
            return new HashSet<string>();
        }

        public SortedDictionary<string, string> GetSortedDictionary(PSObject obj, string name)
        {
            if (!IsHashRefNull(obj, name))
            {
                return (SortedDictionary<string, string>)obj.Properties[name].Value;
            }
            return new SortedDictionary<string, string>();
        }

        public Boolean IsHashRefNull(PSObject obj, string name)
        {
            return obj.Properties[name] == null || obj.Properties[name].Value == null;
        }

        public bool IsSuccess(PSObject obj)
        {
            return GetString(obj, "Result").Equals("Success");
        }
    }
}
