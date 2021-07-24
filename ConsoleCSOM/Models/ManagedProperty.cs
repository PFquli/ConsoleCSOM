using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleCSOM.Models
{
    internal class ManagedProperty
    {
        public string DisplayName { get; set; }
        public string ManagedPropertyName { get; set; }

        public string Value { get; set; }

        public ManagedProperty(string displayName, string managedPropertyName)
        {
            DisplayName = displayName;
            ManagedPropertyName = managedPropertyName;
        }
    }
}