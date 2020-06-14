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
using System.Collections.Generic;
using Polimorfismo.SharePoint.Transaction.Utils;

namespace Polimorfismo.SharePoint.Transaction
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
            get => _fields[key];
            set => _fields[key] = value;
        }

        public IReadOnlyDictionary<string, object> ToDictionary() => _fields;

        #endregion

        #region Constructors / Finalizers

        public SharePointFields(ISharePointMetadata item)
        {
            _fields = item.ToDictionary() ?? new Dictionary<string, object>();

            if (!_fields.ContainsKey(SharePointConstants.FieldNameId))
            {
                _fields.Add(SharePointConstants.FieldNameId, item.Id);
            }
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
