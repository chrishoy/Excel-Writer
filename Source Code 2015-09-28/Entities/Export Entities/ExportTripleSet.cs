using System;

namespace ExcelWriter
{
    internal class ExportTripleSet
    {
        internal ExportTripleSet(IDataPart dataPart, ExportPart exportPart, Template template)
        {
            if (dataPart == null)
            {
                throw new ArgumentNullException("dataPart");
            }
            if (exportPart == null)
            {
                throw new ArgumentNullException("exportPart");
            }
            if (template == null)
            {
                throw new ArgumentNullException("template");
            }

            this.DataPart = dataPart;
            this.Part = exportPart;
            this.Template = template;
        }

        public IDataPart DataPart { get; private set; }
        public ExportPart Part { get; private set; }
        public Template Template { get; private set; }
    }
}
