namespace ExcelWriter
{
    using System;
    using System.Windows.Markup;

    /// <summary>
    /// 
    /// </summary>
    [ContentProperty("Content")]
    public class Sheet : IResource
    {
        #region Private Fields

        private object dataContext;
        private object sheetName;

        #endregion

        public Sheet()
        {
            this.InternalId = Guid.NewGuid().ToString();
        }

        public string Key { get; set; }

        public BaseMap Content { get; set; }

        public object DataContext
        {
            get { return BindingContainer.EvaluateIfRequired(this.dataContext, null); }
            set { this.dataContext = BindingContainer.CreateIfRequired(value); }
        }

        public string PartId { get; set; }  

        public object SheetName
        {
            get { return BindingContainer.EvaluateIfRequired(this.sheetName, this.dataContext); }
            set { this.sheetName = BindingContainer.CreateIfRequired(value); }
        }

        internal string InternalId { get; private set; }        
    }
}
