using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWriter
{
    public static class ObjectExtensions
    {
        public static string GetAssemblyName(this object source)
        {
            return source.GetType().AssemblyQualifiedName;
        }
    }
}
