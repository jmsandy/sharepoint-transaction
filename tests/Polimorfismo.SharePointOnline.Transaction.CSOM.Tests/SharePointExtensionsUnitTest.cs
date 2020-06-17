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
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using Polimorfismo.SharePoint.Transaction;

namespace Polimorfismo.SharePointOnline.Transaction.Tests
{
    /// <summary>
    /// Tests for SharePoint extensions.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-17 07:12:12 PM</Date>
    public class SharePointExtensionsUnitTest : SharePointBaseUnitTest
    {
        [Fact]
        public void SharePoint_Extensions_Fields_Test()
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

            var expectedFieldsDictionary = new Dictionary<string, object>();
            expectedFieldsDictionary.Add("Title", "Title");
            expectedFieldsDictionary.Add("IntegerField", 1);
            expectedFieldsDictionary.Add("ChoiceField", true);
            expectedFieldsDictionary.Add("LookupField", null);
            expectedFieldsDictionary.Add("DecimalField", 22.00M);
            expectedFieldsDictionary.Add("CurrencyField", 1000M);
            expectedFieldsDictionary.Add("OptionField", "Option 2");
            expectedFieldsDictionary.Add("TextField", "Single Line");
            expectedFieldsDictionary.Add("DateField", item.DateField);
            expectedFieldsDictionary.Add("TextArea", "Multiple Lines");
            expectedFieldsDictionary.Add("Created", DateTime.MinValue);
            expectedFieldsDictionary.Add("Modified", DateTime.MinValue);
            expectedFieldsDictionary.Add("PersonOrGroupField", Username);
            expectedFieldsDictionary.Add("LinkField", "https://www.microsoft.com");
            expectedFieldsDictionary.Add("ImageField", "https://www.microsoft.com");

            var fieldsDictionary = item.ToDictionary();

            expectedFieldsDictionary.Count.ShouldEqual(fieldsDictionary.Count);

            fieldsDictionary.Keys.ToList().ForEach(key =>
            {
                expectedFieldsDictionary[key].ShouldEqual(fieldsDictionary[key]);
            });
        }

        [Fact]
        public void SharePoint_Extensions_References_Fields_Test()
        {
            var item = new SharePointListItem
            {
                LookupField = new SharePointAggregatingListItem
                {
                    Id = 1
                }
            };

            var expectedReferencesDictionary = new Dictionary<string, object>();
            expectedReferencesDictionary.Add("LookupField", new SharePointAggregatingListItem { Id = 1 });

            var referencesDictionary = item.GetReferences();

            expectedReferencesDictionary.Count.ShouldEqual(referencesDictionary.Count);

            referencesDictionary.Keys.ToList().ForEach(key =>
            {
                ((SharePointAggregatingListItem)expectedReferencesDictionary[key])
                    .Id.ShouldEqual(((SharePointAggregatingListItem)referencesDictionary[key]).Id);
            });
        }

        [Fact]
        public void SharePoint_Extensions_Users_Fields_Test()
        {
            var item = new SharePointListItem
            {
                PersonOrGroupField = Username
            };

            var expectedUsersDictionary = new Dictionary<string, object>();
            expectedUsersDictionary.Add("PersonOrGroupField", Username);

            var usersDictionary = item.GetUserFields();

            usersDictionary.Keys.ToList().ForEach(key =>
            {
                expectedUsersDictionary[key].ShouldEqual(usersDictionary[key]);
            });
        }

        [Fact]
        public async Task SharePoint_Extensions_Configure_Users_Test()
        {
            var item = new SharePointListItem
            {
                PersonOrGroupField = Username
            };

            var itemTracking = new SharePointItemTracking(item);

            await itemTracking.ConfigureUserFieldsAsync(_sharePointClient);

            itemTracking.Fields["PersonOrGroupField"].ShouldEqual(UserId);
        }

        [Fact]
        public void SharePoint_Extensions_Configure_References_Test()
        {
            var aggregatingListItem = new SharePointAggregatingListItem
            {
                Id = 1,
                TitleField = "Aggregating Item"
            };
            var listItem = new SharePointListItem
            {
                TitleField = "List Item",
                LookupField = aggregatingListItem
            };

            var tracking = new SharePointListItemTracking();
            var itemListTracking = new SharePointItemTracking(listItem);
            var aggregatingListItemTracking = new SharePointItemTracking(aggregatingListItem);

            tracking.Add(itemListTracking);
            tracking.Add(aggregatingListItemTracking);

            itemListTracking.Fields["LookupField"].ShoudBeNull();

            itemListTracking.ConfigureReferences(tracking);

            itemListTracking.Fields["LookupField"].ShouldEqual(aggregatingListItem.Id);
        }

        [Fact]
        public void SharePoint_Extensions_Configure_References_Original_Item_Test()
        {
            var aggregatingListItem = new SharePointAggregatingListItem
            {
                Id = 1,
                TitleField = "Aggregating Item"
            };
            var listItem = new SharePointListItem
            {
                TitleField = "List Item",
                LookupField = aggregatingListItem
            };

            var tracking = new SharePointListItemTracking();
            var itemListTracking = new SharePointItemTracking(null);
            var aggregatingListItemTracking = new SharePointItemTracking(null);

            itemListTracking.LoadOriginalItem(listItem);
            aggregatingListItemTracking.LoadOriginalItem(aggregatingListItem);

            tracking.Add(itemListTracking);
            tracking.Add(aggregatingListItemTracking);

            itemListTracking.OriginalFields["LookupField"].ShoudBeNull();

            itemListTracking.ConfigureReferences(tracking, true);

            itemListTracking.OriginalFields["LookupField"].ShouldEqual(aggregatingListItem.Id);
        }
    }
}
