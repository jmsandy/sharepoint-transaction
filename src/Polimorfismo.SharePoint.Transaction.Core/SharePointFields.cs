using System;
using System.Collections.Generic;

namespace Polimorfismo.SharePoint.Transaction.Core
{
    /// <summary>
    /// Stores the values ​​of the item fields associated with the SharePoint custom list.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 08:35:41 PM</Date>
    internal sealed class SharePointFields : IDisposable
    {
        #region Attributes

        private readonly Dictionary<string, object> _fields;

        #endregion

        #region Properties 

        public int Count => _fields.Count;

        public object this[string key] 
        {
            get =>_fields [key]; 
            set => _fields[key] = value;
        }

        public IReadOnlyDictionary<string, object> ToDictionary() => _fields;

        #endregion

        #region Constructors / Finalizers

        public SharePointFields(ISharePointItem item)
        {
            _fields = item.ToDictionary() ?? new Dictionary<string, object>();
        }

        ~SharePointFields() => Dispose(false);

        #endregion

        #region IDisposable - Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void Dispose(bool disposing)
        {
            if (!disposing) return;
        }

        #endregion
    }
}
