using System;
using System.Linq;
using System.Collections.Generic;

namespace Polimorfismo.SharePoint.Transaction.Core
{
    /// <summary>
    /// Controls the list of operations performed on the custom list to undo them in case of failures.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-05 10:04:26 PM</Date>
    internal sealed class SharePointListItemTracking : IDisposable
    {
        #region Fields

        private List<SharePointItemTracking> _items = new List<SharePointItemTracking>();

        #endregion

        #region Properties

        public IReadOnlyList<SharePointItemTracking> Items => _items;

        #endregion

        #region Constructors / Finalizers

        public SharePointListItemTracking()
        {
        }

        ~SharePointListItemTracking() => Dispose(false);

        #endregion

        #region Methods

        public void Add(SharePointItemTracking item) => _items.Add(item);

        public SharePointItemTracking Get(ISharePointItem item)
            => _items.FirstOrDefault(i => ReferenceEquals(i.Item, item));

        #endregion

        #region IDisposable - Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void Dispose(bool disposing)
        {
            if (disposing)
            {
                _items = null;
            }
        }

        #endregion
    }
}
