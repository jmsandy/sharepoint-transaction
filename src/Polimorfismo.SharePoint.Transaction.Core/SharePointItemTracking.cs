﻿// Copyright 2020 Polimorfismo - José Mauro da Silva Sandy
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
using System.Collections.Generic;
using Polimorfismo.SharePoint.Transaction.Utils;

namespace Polimorfismo.SharePoint.Transaction
{
    /// <summary>
    /// Information about the item present in the customized list during the execution of the commands.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 08:32:26 PM</Date>
    internal sealed class SharePointItemTracking : IDisposable
    {
        #region Properties

        private bool Disposed { get; set; }

        public ISharePointMetadata Item { get; }

        public SharePointFields Fields { get; }

        public SharePointFields OriginalFields { get; private set; }

        public ISharePointMetadata OriginalItem { get; private set; }

        public bool IsOriginalItemLoaded { get; private set; } = false;

        public int Id
        {
            get => (int)Fields[SharePointConstants.FieldNameId];
            set => Fields[SharePointConstants.FieldNameId] = Item.Id = value;
        }

        #endregion

        #region Constructors / Finalizers

        public SharePointItemTracking(ISharePointMetadata item)
        {
            Item = item;
            Fields = new SharePointFields(item);
        }

        #endregion

        #region Methods

        public void LoadOriginalItem(ISharePointMetadata originalItem, Dictionary<string, object> fields = null)
        {
            IsOriginalItemLoaded = true;
            OriginalItem = originalItem;
            OriginalFields = new SharePointFields(originalItem, fields);
        }

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
                Fields?.Dispose();
                OriginalFields?.Dispose();
            }

            Disposed = true;
        }

        #endregion
    }
}
