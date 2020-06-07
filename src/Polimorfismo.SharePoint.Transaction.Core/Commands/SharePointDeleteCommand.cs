using System.Threading.Tasks;

namespace Polimorfismo.SharePoint.Transaction.Core.Commands
{
    /// <summary>
    /// Implements the delete command with the following processes:
    ///     Prepare: obtains the original item;
    ///     Execute: removes the item;
    ///     Undo: inserts the original item in case of failure.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 08:28:58 PM</Date>
    internal class SharePointDeleteCommand<TSharePointItem> : SharePointCommand<TSharePointItem> where TSharePointItem : ISharePointItem
    {
        #region Constructors / Finalizers

        public SharePointDeleteCommand(SharePointClientBase sharePointClient, SharePointItemTracking itemTracking)
            : base(sharePointClient, itemTracking)
        {
        }

        ~SharePointDeleteCommand() => Dispose(false);

        #endregion

        #region ISharePointCommand - Members

        public override async Task Prepare()
        {
            var item = await SharePointClient.GetItemById<TSharePointItem>(SharePointItemTracking.Id);
            SharePointItemTracking.LoadOriginalItem(new SharePointFields(item));
        }

        public override async Task Execute()
        {
            await SharePointClient.DeleteItem<TSharePointItem>(SharePointItemTracking.Id);
        }

        public override async Task Undo()
        {
            await SharePointClient.InsertItem<TSharePointItem>(SharePointItemTracking.OriginalFields.ToDictionary());
        }

        #endregion
    }
}
