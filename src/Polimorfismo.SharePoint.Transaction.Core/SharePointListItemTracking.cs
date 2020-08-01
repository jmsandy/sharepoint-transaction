// Copyright 2020 Polimorfismo - José Mauro da Silva Sandy
// 
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// 
//     http://www.apache.org/licenses/LICENSE-2.0
// 
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

using System;
using System.Linq;
using System.Collections.Generic;

namespace Polimorfismo.SharePoint.Transaction
{
    /// <summary>
    /// Controls the list of operations performed on the custom list to undo them in case of failures.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-05 10:04:26 PM</Date>
    internal sealed class SharePointListItemTracking : IDisposable
    {
        #region Fields

        private readonly List<SharePointItemTracking> _items = new List<SharePointItemTracking>();

        #endregion

        #region Properties

        private bool Disposed { get; set; }

        public IReadOnlyList<SharePointItemTracking> Items => _items;

        #endregion

        #region Constructors / Finalizers

        public SharePointListItemTracking()
        {
        }

        ~SharePointListItemTracking() => Dispose(false);

        #endregion

        #region Methods

        public void Clear()
        {
            _items.Clear();
        }

        public void Add(SharePointItemTracking item) => _items.Add(item);

        public SharePointItemTracking Get(ISharePointItem item, bool isOriginalFields = false)
            => _items.FirstOrDefault(i => ReferenceEquals(isOriginalFields ? i.OriginalItem : i.Item, item));

        #endregion

        #region IDisposable - Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (Disposed) return;

            if (disposing)
            {
                _items.ForEach(item => item.Dispose());
                _items.Clear();
            }

            Disposed = true;
        }

        #endregion
    }
}
