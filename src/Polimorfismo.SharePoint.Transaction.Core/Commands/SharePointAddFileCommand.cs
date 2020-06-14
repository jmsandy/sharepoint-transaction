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
using System.Collections.Generic;

namespace Polimorfismo.SharePoint.Transaction.Commands
{
    /// <summary>
    /// Implements the insert command with the following processes:
    ///     Execute: add the file;
    ///     Undo: remove the inserted file.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-14 08:42:39 AM</Date>
    internal class SharePointAddFileCommand<TSharePointFile> : SharePointCommand<TSharePointFile>
        where TSharePointFile : ISharePointFile, new()
    {
        #region Fields

        private List<string> _createdFolders = null;

        #endregion
        
        #region Constructors / Finalizers

        public SharePointAddFileCommand(SharePointClientBase sharePointClient, SharePointItemTracking itemTracking)
            : base(sharePointClient, itemTracking)
        {
        }

        ~SharePointAddFileCommand() => Dispose(false);

        #endregion

        #region ISharePointCommand - Members

        public override async Task Prepare() => await SharePointItemTracking.ConfigureUserFields(SharePointClient);

        public override async Task Execute()
        {
            var SharePointFile = (ISharePointFile)SharePointItemTracking.Item;

            var fileInfo = await SharePointClient.AddFile<TSharePointFile>(
                SharePointItemTracking.ConfigureReferences(SharePointClient.Tracking), 
                SharePointFile.FileName, SharePointFile.Folder, SharePointFile.InputStream, false);

            SharePointItemTracking.Id = fileInfo.Id;
            _createdFolders = fileInfo.CreatedFolders;
        }

        public override async Task Undo()
        {
            SharePointItemTracking.Id = 0;
            await SharePointClient.RemoveFolders<TSharePointFile>(_createdFolders);
        }

        #endregion
    }
}
