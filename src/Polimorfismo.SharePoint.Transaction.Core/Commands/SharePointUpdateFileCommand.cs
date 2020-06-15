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

using System.Threading.Tasks;

namespace Polimorfismo.SharePoint.Transaction.Commands
{
    /// <summary>
    /// Implements the update command with the following processes:
    ///     Prepare: obtains the original file;
    ///     Execute: update the file;
    ///     Undo: update the original file in case of failure.    
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-15 07:19:12 PM</Date>
    internal class SharePointUpdateFileCommand<TSharePointFile> : SharePointCommand<TSharePointFile>
        where TSharePointFile : ISharePointFile, new()
    {
        #region Constructors / Finalizers

        public SharePointUpdateFileCommand(SharePointClientBase sharePointClient, SharePointItemTracking itemTracking)
            : base(sharePointClient, itemTracking)
        {
        }

        ~SharePointUpdateFileCommand() => Dispose(false);

        #endregion

        #region ISharePointCommand - Members

        public override async Task Prepare()
        {
            await SharePointItemTracking.ConfigureUserFields(SharePointClient);

            var item = await SharePointClient.GetFileById<TSharePointFile>(SharePointItemTracking.Id);
            SharePointItemTracking.LoadOriginalItem(item);
        }

        public override async Task Execute()
        {
            var sharePointFile = (ISharePointFile)SharePointItemTracking.Item;

            await SharePointClient.AddFile<TSharePointFile>(
                SharePointItemTracking.ConfigureReferences(SharePointClient.Tracking),
                sharePointFile.FileName, sharePointFile.Folder, sharePointFile.InputStream, true);
        }

        public override async Task Undo()
        {
            var sharePointFile = (ISharePointFile)SharePointItemTracking.OriginalItem;

            await SharePointClient.AddFile<TSharePointFile>(
                SharePointItemTracking.ConfigureReferences(SharePointClient.Tracking, true),
                sharePointFile.FileName, sharePointFile.Folder, sharePointFile.InputStream, true);
        }

        #endregion
    }
}
