using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace Owl.Word
{
    public interface ICellProperties
    {
        JustificationValues Justification { get; set; }

        bool DrawShade { get; set; }
        Drawing.Color ShadeColor { get; set; }


        bool DrawBorder { get; set; }
        Drawing.Color BorderColor { get; set; }
        /// <summary>
        /// Specifies the width of the border. Table borders are line borders, and so the width is specified in eighths of a point, with a minimum value of two (1/4 of a point) and a maximum value of 96 (twelve points)
        /// </summary>
        uint BorderSize { get; set; }


        Action<TableCell, Data.DataColumn, string> OnCellRender { get; set; }
    }
}
