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
using CamlexNET;
using System.Linq;
using System.Threading.Tasks;
using Polimorfismo.SharePoint.Transaction.Utils;

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
                ChoiceField = true,
                TitleField = "Title",
                DecimalField = 22.00M,
                CurrencyField = 1000M,
                OptionField = "Option 2",
                TextField = "Single Line",
                TextArea = "Multiple Lines",
                PersonOrGroupField = Username,
                DateField = DateTime.Now.Date,
                LinkField = "https://www.microsoft.com",
                ImageField = "https://www.microsoft.com"
            };

            _sharePointClient.AddItem(item);

            await _sharePointClient.SaveChanges();

            var expectedItem = await _sharePointClient.GetItemById<SharePointListItem>(item.Id);

            Assert.NotNull(expectedItem);
            item.Id.ShouldEqual(expectedItem.Id);
            item.TextArea.ShouldEqual(expectedItem.TextArea);
            item.LinkField.ShouldEqual(expectedItem.LinkField);
            item.TextField.ShouldEqual(expectedItem.TextField);
            item.TitleField.ShouldEqual(expectedItem.TitleField);
            item.ImageField.ShouldEqual(expectedItem.ImageField);
            item.ChoiceField.ShouldEqual(expectedItem.ChoiceField);
            item.OptionField.ShouldEqual(expectedItem.OptionField);
            item.DecimalField.ShouldEqual(expectedItem.DecimalField);
            item.IntegerField.ShouldEqual(expectedItem.IntegerField);
            item.CurrencyField.ShouldEqual(expectedItem.CurrencyField);
            item.DateField.ShouldEqual(expectedItem.DateField.Value.Date);
        }

        [Fact]
        public async Task SharePoint_GetItems_Success_Test()
        {
            var item = new SharePointListItem
            {
                TitleField = Guid.NewGuid().ToString()
            };

            _sharePointClient.AddItem(item);

            await _sharePointClient.SaveChanges();

            var items = await _sharePointClient.GetItems<SharePointListItem>(
                Camlex.Query().Where(i => (string)i[SharePointConstants.FieldNameTitle] == item.TitleField).ToCamlQuery());

            item.TitleField.ShouldEqual(items.Single().TitleField);
        }

        [Fact]
        public async Task SharePoint_UpdateItem_Success_Test()
        {
            var item = new SharePointListItem
            {
                TitleField = Guid.NewGuid().ToString()
            };

            _sharePointClient.AddItem(item);
            await _sharePointClient.SaveChanges();

            var updatedItem = await _sharePointClient.GetItemById<SharePointListItem>(item.Id);

            var updatedTitle = $"Updated - {item.TitleField}";
            updatedItem.TitleField = updatedTitle;

            _sharePointClient.UpdateItem(updatedItem);
            await _sharePointClient.SaveChanges();

            var expectedItem = await _sharePointClient.GetItemById<SharePointListItem>(updatedItem.Id);
            updatedItem.TitleField = updatedTitle;

            updatedTitle.ShouldEqual(expectedItem.TitleField);
        }

        [Fact]
        public async Task SharePoint_DeleteItem_Success_Test()
        {
            var item = new SharePointListItem
            {
                TitleField = Guid.NewGuid().ToString()
            };

            _sharePointClient.AddItem(item);
            await _sharePointClient.SaveChanges();

            var deletedItem = await _sharePointClient.GetItemById<SharePointListItem>(item.Id);

            _sharePointClient.DeleteItem(deletedItem);
            await _sharePointClient.SaveChanges();

            var expectedItem = await _sharePointClient.GetItemById<SharePointListItem>(item.Id);

            Assert.Null(expectedItem);
        }

        [Fact]
        public async Task SharePoint_AddItem_Association_Unknown_Id_Success_Test()
        {
            var aggregatingListItem = new SharePointAggregatingListItem
            {
                TitleField = "Aggregating Item"
            };
            var listItem = new SharePointListItem
            {
                TitleField = "List Item",
                LookupField = aggregatingListItem
            };

            _sharePointClient.AddItem(listItem);
            _sharePointClient.AddItem(aggregatingListItem);

            await _sharePointClient.SaveChanges();

            var expectedListItem = await _sharePointClient.GetItemById<SharePointListItem>(listItem.Id);

            expectedListItem.TitleField.ShouldEqual(listItem.TitleField);
            expectedListItem.LookupFieldId.ShouldEqual(aggregatingListItem.Id);
        }

        [Fact]
        public async Task SharePoint_AddItem_Association_Known_Id_Success_Test()
        {
            var aggregatingListItem = new SharePointAggregatingListItem
            {
                TitleField = "Aggregating Item"
            };

            _sharePointClient.AddItem(aggregatingListItem);

            await _sharePointClient.SaveChanges();

            var listItem = new SharePointListItem
            {
                TitleField = "List Item 1",
                LookupFieldId = aggregatingListItem.Id
            };
            var listItem2 = new SharePointListItem
            {
                TitleField = "List Item 2",
                LookupFieldId = aggregatingListItem.Id
            };

            _sharePointClient.AddItem(listItem);
            _sharePointClient.AddItem(listItem2);

            await _sharePointClient.SaveChanges();

            var items = await _sharePointClient.GetItems<SharePointListItem>(
                Camlex.Query().Where(i => i["LookupField"] == (DataTypes.LookupId)aggregatingListItem.Id.ToString()).ToCamlQuery());

            items.Count.ShouldEqual(2);
        }
    }
}
