using System;
using System.Collections.Generic;
using System.Text;

namespace Owl.Drawing
{
    class ColorTranslator
    {
        public static string ToHtml(Color c)
        {
            string colorString = String.Empty;

            if (c.IsEmpty)
                return colorString;

            if (c.IsSystemColor)
            {
                switch (c.ToKnownColor())
                {
                    case KnownColor.ActiveBorder: colorString = "activeborder"; break;
                    case KnownColor.GradientActiveCaption:
                    case KnownColor.ActiveCaption: colorString = "activecaption"; break;
                    case KnownColor.AppWorkspace: colorString = "appworkspace"; break;
                    case KnownColor.Desktop: colorString = "background"; break;
                    case KnownColor.Control: colorString = "buttonface"; break;
                    case KnownColor.ControlLight: colorString = "buttonface"; break;
                    case KnownColor.ControlDark: colorString = "buttonshadow"; break;
                    case KnownColor.ControlText: colorString = "buttontext"; break;
                    case KnownColor.ActiveCaptionText: colorString = "captiontext"; break;
                    case KnownColor.GrayText: colorString = "graytext"; break;
                    case KnownColor.HotTrack:
                    case KnownColor.Highlight: colorString = "highlight"; break;
                    case KnownColor.MenuHighlight:
                    case KnownColor.HighlightText: colorString = "highlighttext"; break;
                    case KnownColor.InactiveBorder: colorString = "inactiveborder"; break;
                    case KnownColor.GradientInactiveCaption:
                    case KnownColor.InactiveCaption: colorString = "inactivecaption"; break;
                    case KnownColor.InactiveCaptionText: colorString = "inactivecaptiontext"; break;
                    case KnownColor.Info: colorString = "infobackground"; break;
                    case KnownColor.InfoText: colorString = "infotext"; break;
                    case KnownColor.MenuBar:
                    case KnownColor.Menu: colorString = "menu"; break;
                    case KnownColor.MenuText: colorString = "menutext"; break;
                    case KnownColor.ScrollBar: colorString = "scrollbar"; break;
                    case KnownColor.ControlDarkDark: colorString = "threeddarkshadow"; break;
                    case KnownColor.ControlLightLight: colorString = "buttonhighlight"; break;
                    case KnownColor.Window: colorString = "window"; break;
                    case KnownColor.WindowFrame: colorString = "windowframe"; break;
                    case KnownColor.WindowText: colorString = "windowtext"; break;
                }
            }
            else if (c.IsNamedColor && !c.IsKnownColor)
            {
                if (c == Color.LightGray)
                {
                    // special case due to mismatch between Html and enum spelling
                    colorString = "LightGrey";
                }
                else
                {
                    colorString = c.Name;
                }
            }
            else
            {
                colorString = "#" + c.R.ToString("X2", null) +
                                    c.G.ToString("X2", null) +
                                    c.B.ToString("X2", null);
            }

            return colorString;
        }

    }
}
