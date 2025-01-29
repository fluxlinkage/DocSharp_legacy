using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// NOTE: This record is called SXIDSTM in the old version of the specification
    /// </summary>
    [BiffRecord(RecordType.SXStreamID)] 
    public class SXStreamID : BiffRecord
    {
        public const RecordType ID = RecordType.SXStreamID;

        public SXStreamID(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            // TODO: place code here
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
