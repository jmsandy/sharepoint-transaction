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
    /// Implements the insert command with the following processes:
    ///     Execute: add the item;
    ///     Undo: remove the inserted item.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 08:26:20 PM</Date>
    internal class SharePointAddItemCommand<TSharePointItem> : SharePointCommand<TSharePointItem>
        where TSharePointItem : ISharePointItem, new()
    {
        #region Constructors / Finalizers

        public SharePointAddItemCommand(SharePointClientBase sharePointClient, SharePointItemTracking itemTracking)
            : base(sharePointClient, itemTracking)
        {
        }

        ~SharePointAddItemCommand() => Dispose(false);

        #endregion

        #region ISharePointCommand - Members

        public override async Task Prepare() => await SharePointItemTracking.ConfigureUserFields(SharePointClient);

        public override async Task Execute()
        {
            var id = await SharePointClient.AddItem<TSharePointItem>(
                SharePointItemTracking.ConfigureReferences(SharePointClient.Tracking));

            SharePointItemTracking.Id = id;
        }

        public override async Task Undo()
        {
            await SharePointClient.DeleteItem<TSharePointItem>(SharePointItemTracking.Id);

            SharePointItemTracking.Id = 0;
        }

        #endregion
    }
}