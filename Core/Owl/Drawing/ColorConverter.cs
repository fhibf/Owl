using System.Reflection;
using System.Collections.Generic;
using System;
using System.Linq;

namespace Owl.Drawing
{
    /// <include file='doc\ColorConverter.uex' path='docs/doc[@for="ColorConverter"]/*' />
    /// <devdoc>
    ///      ColorConverter is a class that can be used to convert
    ///      colors from one data type to another.  Access this
    ///      class through the TypeDescriptor.
    /// </devdoc>
    public class ColorConverter 
    {
        private static string ColorConstantsLock = "colorConstants";
        private static Dictionary<string, Color> colorConstants;
        private static string SystemColorConstantsLock = "systemColorConstants";
        private static Dictionary<string, Color> systemColorConstants;

        internal static object GetNamedColor(string name)
        {
            object color = null;
            // First, check to see if this is a standard name.
            //
            color = Colors[name];
            if (color != null)
            {
                return color;
            }
            // Ok, how about a system color?
            //
            color = SystemColors[name];
            return color;
        }

        private static Dictionary<string, Color> SystemColors
        {
            get
            {
                if (systemColorConstants == null)
                {
                    lock (SystemColorConstantsLock)
                    {
                        if (systemColorConstants == null)
                        {
                            Dictionary<string, Color> tempHash = new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase);
                            FillConstants(tempHash, typeof(SystemColors));
                            systemColorConstants = tempHash;
                        }
                    }
                }

                return systemColorConstants;
            }
        }


        private static IDictionary<string,Color> Colors
        {
            get
            {
                if (colorConstants == null)
                {
                    lock (ColorConstantsLock)
                    {
                        if (colorConstants == null)
                        {
                            Dictionary<string, Color> tempHash = new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase);
                            FillConstants(tempHash, typeof(Color));
                            colorConstants = tempHash;
                        }
                    }
                }

                return colorConstants;
            }
        }

        private static void FillConstants(Dictionary<string, Color> hash, Type enumType)
        {
            MethodAttributes attrs = MethodAttributes.Public | MethodAttributes.Static;
            PropertyInfo[] props = enumType.GetTypeInfo().DeclaredProperties.ToArray();

            for (int i = 0; i < props.Length; i++)
            {
                PropertyInfo prop = props[i];
                if (prop.PropertyType == typeof(Color))
                {
                    MethodInfo method = prop.GetMethod;
                    if (method != null && (method.Attributes & attrs) == attrs)
                    {
                        object[] tempIndex = null;
                        hash[prop.Name] = (Color)prop.GetValue(null, tempIndex);
                    }
                }
            }
        }
    }
}
