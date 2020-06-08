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

using System.Threading.Tasks;

namespace Polimorfismo.SharePoint.Transaction.Commands
{
    /// <summary>
    /// Implements the delete command with the following processes:
    ///     Prepare: obtains the original item;
    ///     Execute: removes the item;
    ///     Undo: inserts the original item in case of failure.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 08:28:58 PM</Date>
    internal class SharePointDeleteCommand<TSharePointItem> : SharePointCommand<TSharePointItem> 
        where TSharePointItem : ISharePointItem, new()
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
