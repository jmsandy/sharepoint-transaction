using Polimorfismo.SharePoint.Transaction.Core.Utils;

namespace Polimorfismo.SharePoint.Transaction.Core
{
    /// <summary>
    /// Information about the item present in the customized list during the execution of the commands.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 08:32:26 PM</Date>
    internal class SharePointItemTracking
    {
        #region Properties

        public ISharePointItem Item { get; }

        public SharePointFields Fields { get; }

        public SharePointFields OriginalFields { get; private set; }

        public bool IsOriginalItemLoaded { get; private set; } = false;

        public int Id
        {
            get => (int)Fields[SharePointConstants.FieldNameId];
            set => Fields[SharePointConstants.FieldNameId] = value;
        }

        #endregion

        #region Constructors / Finalizers

        public SharePointItemTracking(ISharePointItem item)
        {
            Item = item;
            Fields = new SharePointFields(item);
        }

        #endregion

        #region Methods

        public void LoadOriginalItem(SharePointFields originalItem)
        {
            IsOriginalItemLoaded = true;
            OriginalFields = originalItem;
        }

        #endregion
    }
}
