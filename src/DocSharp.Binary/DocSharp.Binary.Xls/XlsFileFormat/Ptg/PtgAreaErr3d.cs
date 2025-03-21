﻿using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgAreaErr3d : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgAreaErr3d;
        public ushort ixti;

        public PtgAreaErr3d(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 11;
            this.Data = "#REF!";
            this.type = PtgType.Operand;
            this.ixti = reader.ReadUInt16(); 
            reader.ReadBytes(8);             
        }
    }
}
