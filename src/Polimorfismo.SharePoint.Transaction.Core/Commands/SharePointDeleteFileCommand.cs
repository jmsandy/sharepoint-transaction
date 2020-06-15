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
    /// Implements the delete command with the following processes:
    ///     Prepare: obtains the original file;
    ///     Execute: remove the file;
    ///     Undo: insert the original file in case of failure.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-14 09:20:49 PM</Date>
    internal class SharePointDeleteFileCommand<TSharePointFile> : SharePointCommand<TSharePointFile>
        where TSharePointFile : ISharePointFile, new()
    {
        #region Constructors / Finalizers

        public SharePointDeleteFileCommand(SharePointClientBase sharePointClient, SharePointItemTracking itemTracking)
            : base(sharePointClient, itemTracking)
        {
        }

        ~SharePointDeleteFileCommand() => Dispose(false);

        #endregion

        #region ISharePointCommand - Members

        public override async Task Prepare()
        {
            var item = await SharePointClient.GetFileById<TSharePointFile>(SharePointItemTracking.Id);
            SharePointItemTracking.LoadOriginalItem(item);
        }

        public override async Task Execute()
        {
            await SharePointClient.DeleteFile<TSharePointFile>(SharePointItemTracking.Id);
        }

        public override async Task Undo()
        {
            var sharePointFile = (ISharePointFile)SharePointItemTracking.OriginalItem;

            await SharePointClient.AddFile<TSharePointFile>(
                SharePointItemTracking.OriginalFields.ToDictionary(),
                sharePointFile.FileName, sharePointFile.Folder, sharePointFile.InputStream, false);
        }

        #endregion
    }
}
