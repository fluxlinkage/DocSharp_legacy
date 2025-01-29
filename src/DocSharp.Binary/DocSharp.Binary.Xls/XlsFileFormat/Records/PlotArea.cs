﻿using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This empty record specifies that the Frame record that immediately follows this 
    /// record specifies properties of the plot area.
    /// </summary>
    [BiffRecord(RecordType.PlotArea)]
    public class PlotArea : BiffRecord
    {
        public const RecordType ID = RecordType.PlotArea;

        public PlotArea(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // Record is emty

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
