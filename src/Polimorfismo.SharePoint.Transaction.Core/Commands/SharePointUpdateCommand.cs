using System.Threading.Tasks;

namespace Polimorfismo.SharePoint.Transaction.Core.Commands
{
    /// <summary>
    /// Implements the update command with the following processes:
    ///     Prepare: obtains the original item;
    ///     Execute: updates the item;
    ///     Undo: update the original item in case of failure.    
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 08:27:45 PM</Date>
    internal class SharePointUpdateCommand<TSharePointItem> : SharePointCommand<TSharePointItem> where TSharePointItem : ISharePointItem
    {
        #region Constructors / Finalizers

        public SharePointUpdateCommand(SharePointClientBase sharePointClient, SharePointItemTracking itemTracking)
            : base(sharePointClient, itemTracking)
        {
        }

        ~SharePointUpdateCommand() => Dispose(false);

        #endregion

        #region ISharePointCommand - Members

        public override async Task Prepare()
        {
            var item = await SharePointClient.GetItemById<TSharePointItem>(SharePointItemTracking.Id);
            SharePointItemTracking.LoadOriginalItem(new SharePointFields(item));
        }

        public override async Task Execute()
        {
            await SharePointClient.UpdateItem<TSharePointItem>(SharePointItemTracking.Id, SharePointItemTracking.Fields.ToDictionary());
        }

        public override async Task Undo()
        {
            await SharePointClient.UpdateItem<TSharePointItem>(SharePointItemTracking.Id, SharePointItemTracking.Fields.ToDictionary());
        }

        #endregion
    }
}
