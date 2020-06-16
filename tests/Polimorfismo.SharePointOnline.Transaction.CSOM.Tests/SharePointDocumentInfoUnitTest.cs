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

using Xunit;
using System.Collections.Generic;
using Polimorfismo.SharePoint.Transaction;
using Polimorfismo.SharePoint.Transaction.Resources;

namespace Polimorfismo.SharePointOnline.Transaction.Tests
{
    /// <summary>
    /// Tests for document info.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-15 09:37:11 PM</Date>
    public class SharePointDocumentInfoUnitTest : SharePointBaseUnitTest
    {
        [Fact]
        public void SharePoint_DocumentInfo_Add_Documents_Success_Test()
        {
            var folder = new SharePointDocumentInfo(0, "Folder", null, false);
            folder.AddDocuments(new List<SharePointDocumentInfo>()
            {
                new SharePointDocumentInfo(1, "File1.txt", null, true),
                new SharePointDocumentInfo(2, "File2.txt", null, true)
            });

            folder.Documents.Count.ShouldEqual(2);
        }

        [Fact]
        public void SharePoint_DocumentInfo_Add_Documents_Failure_Only_Folders_Can_Receive_Documents_Test()
        {
            var folder = new SharePointDocumentInfo(0, "Folder", null, true);
            var exception = Assert.Throws<SharePointException>(() =>
            {
                folder.AddDocuments(new List<SharePointDocumentInfo>()
                {
                    new SharePointDocumentInfo(1, "File1.txt", null, true),
                    new SharePointDocumentInfo(2, "File2.txt", null, true)
                });
            });

            exception.ErrorCode.ShouldEqual(SharePointErrorCode.OnlyFoldersCanReceiveDocuments);
            exception.Message.ShouldEqual(SharePointMessages.ERR402);
        }
    }
}
