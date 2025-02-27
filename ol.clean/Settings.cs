using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace ol.clean
{
    [XmlRootAttribute("Settings", Namespace = "http://sample.tusharkarsan.co.uk/ol.clean", IsNullable = false)]
    public class Settings
    {
        [XmlAttribute]
        public string version = "1.0";

        [XmlArrayItem("domain")]
        public string[] ExcludeDomain = {
            "aol.com",
            "bellsouth.net",
            "btinternet.com",
            "charter.net",
            "comcast.net",
            "cox.net",
            "earthlink.net",
            "gmail.co.uk",
            "gmail.com",
            "google.co.uk",
            "google.com",
            "hotmail.co.uk",
            "hotmail.com",
            "live.co.uk",
            "live.com",
            "microsoft.co.uk",
            "microsoft.com",
            "msn.co.uk",
            "msn.com",
            "outlook.co.uk",
            "outlook.com",
            "ntlworld.com",
            "rediffmail.com",
            "sbcglobal.net",
            "shaw.ca",
            "verizon.net",
            "yahoo.ca",
            "yahoo.co.in",
            "yahoo.co.uk",
            "yahoo.com"
        };

        [XmlArrayAttribute("Rules")]
        public TypeRule[] rules = { };

        /// <summary>
        /// The name for the settings is fixed.
        /// </summary>
        /// <returns>Returns user profile specific settings filename.</returns>
        internal static string GetFilename()
        {
            var root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(root, "ol.clean.xml");
        }

        /// <summary>
        /// Reads config from user profile area.
        /// </summary>
        /// <returns></returns>
        internal static Settings Load()
        {
            var filename = Settings.GetFilename();

            if (!File.Exists(filename))
                return new Settings();

            using (var stream = File.OpenRead(filename))
            {
                var serializer = new XmlSerializer(typeof(Settings));
                return serializer.Deserialize(stream) as Settings;
            }
        }

        /// <summary>
        /// Writes this config to user profile area.
        /// </summary>
        internal void Save()
        {
            using (var writer = new StreamWriter(Settings.GetFilename(), false, System.Text.Encoding.UTF8))
            {
                var serializer = new XmlSerializer(this.GetType());
                serializer.Serialize(writer, this);
                writer.Flush();
            }
        }

        internal TypeRule FindRule(string address)
        {
            address = address.ToLowerInvariant();

            foreach (var rule in rules)
            {
                if (rule.EndsWith)
                {
                    if (address.Length > rule.Criteria.Length)
                        if (string.Compare(rule.Criteria, address.Substring(address.Length - rule.Criteria.Length), StringComparison.InvariantCultureIgnoreCase) == 0)
                            return rule;
                }
                else
                {
                    if (string.Compare(rule.Criteria, address, StringComparison.InvariantCultureIgnoreCase) == 0)
                        return rule;
                }
            }

            return null;
        }

        internal TypeRule Find(string address, bool findExact)
        {
            var index = address.IndexOf('@');
            var useAddress = (findExact == false && index > 0) ? address.Substring(index) : address;

            foreach (var rule in rules)
            {
                if (string.Compare(rule.Criteria,useAddress, StringComparison.InvariantCultureIgnoreCase) == 0)
                    return rule;
            }

            return null;
        }

        internal void Add(string address, bool addExact, int period)
        {
            var index = address.IndexOf('@');
            var useAddress = (addExact == false && index > 0) ? address.Substring(index) : address;
            var newRule = new TypeRule() { Criteria = useAddress, EndsWith = !addExact, Period = period };

            var oldList = new List<TypeRule>(rules);
            oldList.Add(newRule);
            oldList.Sort(delegate(TypeRule r1, TypeRule r2)
            {
                int compareVal = r1.Criteria.CompareTo(r2.Criteria);
                if (compareVal == 0)
                    compareVal = r1.Period.CompareTo(r2.Period);
                return compareVal;
            });
            rules = oldList.ToArray();
        }

        internal bool IsExcluded(string address)
        {
            foreach (var criteria in ExcludeDomain)
            {
                if (address.Length > criteria.Length)
                    if (string.Compare(criteria, address.Substring(address.Length - criteria.Length), StringComparison.InvariantCultureIgnoreCase) == 0)
                        return true;
            }
            return false;
        }
    }

    [XmlType("Rule")]
    public class TypeRule
    {
        /// <summary>
        /// Search criteria, Exact match or EndsWith string value.
        /// </summary>
        public string Criteria = string.Empty;

        /// <summary>
        /// Otherwise perform 'Exact' match if set to True.
        /// </summary>
        public bool EndsWith = false;

        /// <summary>
        /// Internal falg to indicate length in days or months.
        /// </summary>
        public int Period = 0;

        public override string ToString()
        {
            return string.Format("{0} - {1} - {2}", Criteria, Period.ToString("000"), EndsWith);
        }
    }
}
