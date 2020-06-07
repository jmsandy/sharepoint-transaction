using System.Threading.Tasks;

namespace Polimorfismo.SharePoint.Transaction.Core.Commands
{
    /// <summary>
    /// Implements the insert command with the following processes:
    ///     Execute: inserts the item;
    ///     Undo: removes the inserted item.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 08:26:20 PM</Date>
    internal class SharePointInsertCommand<TSharePointItem> : SharePointCommand<TSharePointItem> where TSharePointItem : ISharePointItem
    {
        #region Constructors / Finalizers

        public SharePointInsertCommand(SharePointClientBase sharePointClient, SharePointItemTracking itemTracking)
            : base(sharePointClient, itemTracking)
        {
        }

        ~SharePointInsertCommand() => Dispose(false);

        #endregion

        #region ISharePointCommand - Members

        public override Task Prepare() => Task.CompletedTask;

        public override async Task Execute()
        {
            var id = await SharePointClient.InsertItem<TSharePointItem>(
                SharePointItemTracking.ConfigureReferences(SharePointClient.Tracking));

            SharePointItemTracking.Id = id;
        }

        public override async Task Undo()
        {
            await SharePointClient.DeleteItem<TSharePointItem>(SharePointItemTracking.Id);
        }

        #endregion
    }
}
