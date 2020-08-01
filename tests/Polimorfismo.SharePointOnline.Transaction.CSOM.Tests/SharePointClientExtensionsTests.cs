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
using System;
using FluentAssertions;
using Microsoft.SharePoint.Client;
using Polimorfismo.SharePoint.Transaction;
using Polimorfismo.SharePoint.Transaction.Commons.Tests;

namespace Polimorfismo.SharePointOnline.Transaction.Tests
{
    /// <summary>
    /// Testing extensions operations.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-07-30 05:37:55 PM</Date>
    public class SharePointClientExtensionsTests
    {
        [Trait("Category", "SharePointOnline - Extensions")]
        [Fact(DisplayName = "Exception when triggering the copy of a null listItem")]
        public void SharePointClientExtensions_CopyListItemTo_ListItemNull()
        {
            // Arrange 
            var clientContext = new ClientContext("http://localhost");

            // Act & Assert
            var exception = Assert.Throws<ArgumentNullException>(() => clientContext.CopyListItemTo<SharePointListItem>(null));
            exception.Message.Should().Contain("listItem");
        }

        [Trait("Category", "SharePointOnline - Extensions")]
        [Theory(DisplayName = "Exception when getting the list when the name is null or empty")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("    ")]
        public void SharePointClientExtensions_GetList_ListNameNullOrEmpty(string listName)
        {
            // Arrange 
            var clientContext = new ClientContext("http://localhost");

            // Act & Assert
            var exception = Assert.Throws<ArgumentNullException>(() => clientContext.GetList(listName));
            exception.Message.Should().Contain("listName");
        }
    }
}