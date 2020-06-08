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
using System.Threading.Tasks;

namespace Polimorfismo.SharePointOnline.Transaction.Tests
{
    /// <summary>
    /// Testing operations to insert, update, remove and retrieve items in a SharePoint custom list.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-07 09:36:53 PM</Date>
    public class SharePointCrudUnitTest : SharePointBaseUnitTest
    {
        [Fact]
        public async Task SharePoint_AddItem_Success_Test()
        {
            var item = new SharePointListItem
            {
                IntegerField = 1,
                TitleField = "Title",
                DecimalField = 22.00M,
                OptionField = "Option 2",
                TextField = "Single Line",
                TextArea = "Multiple Lines"
            };

            _sharePointClient.AddItem(item);

            await _sharePointClient.SaveChanges();

            var expectedItem = await _sharePointClient.GetItemById<SharePointListItem>(item.Id);

            Assert.NotNull(expectedItem);
            item.Id.ShouldEqual(expectedItem.Id);
            item.TextArea.ShouldEqual(expectedItem.TextArea);
            item.TextField.ShouldEqual(expectedItem.TextField);
            item.TitleField.ShouldEqual(expectedItem.TitleField);
            item.OptionField.ShouldEqual(expectedItem.OptionField);
            item.DecimalField.ShouldEqual(expectedItem.DecimalField);
            item.IntegerField.ShouldEqual(expectedItem.IntegerField);
        }
    }
}
