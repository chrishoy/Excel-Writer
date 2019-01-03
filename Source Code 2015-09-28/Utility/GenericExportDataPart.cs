namespace ExcelWriter
{
    /// <summary>
    /// Generic Data Part for consumption by ExcelWriter<br/>
    /// </summary>
    public class GenericExportDataPart : IPreparableDataPart
    {
        #region Constructors

        /// <summary>
        /// Initialises a new instance of the <see cref="GenericExportDataPart"/> class.
        /// </summary>
        /// <param name="sourceDataPart">The source data and identifier which will be rendered as an Exel document.</param>
        public GenericExportDataPart(ExportDataPart sourceDataPart)
        {
            this.PartId = sourceDataPart.PartId;
            this.Data = sourceDataPart.Data;
        }

        #endregion Construction
    
        public void Prepare(object parent, IExcelPreparable elementToPrepare)
        {
            throw new System.NotImplementedException();
        }

        /// <summary>
        /// Gets the source data for this <see cref="GenericExportDataPart"/>
        /// </summary>
        public object Data { get; private set; }

        public string PartId { get; set; }

        /// <summary>
        /// Irrelivant interface member.
        /// TODO: Get rid of this...
        /// </summary>
        public int RowCount
        {
            get { throw new System.NotImplementedException(); }
        }
    }
}
