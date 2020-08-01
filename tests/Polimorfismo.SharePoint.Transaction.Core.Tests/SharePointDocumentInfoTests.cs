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
using FluentAssertions;
using System.Collections.Generic;
using Polimorfismo.SharePoint.Transaction.Resources;

namespace Polimorfismo.SharePoint.Transaction.Core.Tests
{
    /// <summary>
    /// Tests for document info.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-15 09:37:11 PM</Date>
    public class SharePointDocumentInfoTests
    {
        [Trait("Category", "SharePointCore - DocumentInfo")]
        [Fact(DisplayName = "Adding document in SharePoint folder with success")]
        public void SharePointDocumentInfo_AddDocument_Success()
        {
            // Arrange
            var folder = new SharePointDocumentInfo(0, "Folder", null, false);

            // Act
            folder.AddDocument(new SharePointDocumentInfo(1, "File1.txt", null, true));

            // Assert
            folder.Documents.Count.Should().Be(1);
        }

        [Trait("Category", "SharePointCore - DocumentInfo")]
        [Fact(DisplayName = "Adding document error. Only folders can receive documents")]
        public void SharePointDocumentInfo_AddDocument_OnlyFoldersCanReceiveDocuments()
        {
            // Arrange
            var document = new SharePointDocumentInfo(1, "File1.txt", null, true);

            // Act
            var exception = Assert.Throws<SharePointException>(() =>
            {
                document.AddDocument(new SharePointDocumentInfo(1, "File1.txt", null, true));
            });

            // Assert
            exception.Message.Should().Be(SharePointMessages.ERR402);
            exception.ErrorCode.Should().Be(SharePointErrorCode.OnlyFoldersCanReceiveDocuments);
        }

        [Trait("Category", "SharePointCore - DocumentInfo")]
        [Fact(DisplayName = "Adding documents in SharePoint folder with success")]
        public void SharePointDocumentInfo_AddDocuments_Success()
        {
            // Arrange
            var folder = new SharePointDocumentInfo(0, "Folder", null, false);

            // Act
            folder.AddDocuments(new List<SharePointDocumentInfo>()
            {
                new SharePointDocumentInfo(1, "File1.txt", null, true),
                new SharePointDocumentInfo(2, "File2.txt", null, true)
            });

            // Assert
            folder.Documents.Count.Should().Be(2);
        }

        [Trait("Category", "SharePointCore - DocumentInfo")]
        [Fact(DisplayName = "Adding documents error. Only folders can receive documents")]
        public void SharePointDocumentInfo_AddDocuments_OnlyFoldersCanReceiveDocuments()
        {
            // Arrange
            var folder = new SharePointDocumentInfo(0, "Folder", null, true);

            // Act
            var exception = Assert.Throws<SharePointException>(() =>
            {
                folder.AddDocuments(new List<SharePointDocumentInfo>()
                {
                    new SharePointDocumentInfo(1, "File1.txt", null, true),
                    new SharePointDocumentInfo(2, "File2.txt", null, true)
                });
            });

            // Assert
            exception.Message.Should().Be(SharePointMessages.ERR402);
            exception.ErrorCode.Should().Be(SharePointErrorCode.OnlyFoldersCanReceiveDocuments);
        }
    }
}
