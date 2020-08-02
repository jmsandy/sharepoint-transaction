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
using FluentAssertions;
using System.Collections.Generic;
using Polimorfismo.SharePoint.Transaction;
using Polimorfismo.SharePoint.Transaction.Commons.Tests;

namespace Polimorfismo.SharePointOnline.Transaction.Tests
{
    /// <summary>
    /// Tests for SharePoint extensions.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-17 07:12:12 PM</Date>
    public class SharePointExtensionsTests
    {
        [Trait("Category", "SharePointCore - Extensions")]
        [Fact(DisplayName = "Converts SharePoint object to Dictionary with success")]
        public void SharePointExtensions_ToDictionary_ConvertsObjectToDictionaryWithSuccess()
        {
            // Arrange
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
                DateField = DateTime.Now.Date,
                PersonOrGroupField = "Username",
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
            expectedFieldsDictionary.Add("PersonOrGroupField", "Username");
            expectedFieldsDictionary.Add("LinkField", "https://www.microsoft.com");
            expectedFieldsDictionary.Add("ImageField", "https://www.microsoft.com");

            // Act
            var fieldsDictionary = item.ToDictionary();

            // Assert
            expectedFieldsDictionary.Count.Should().Be(fieldsDictionary.Count);

            fieldsDictionary.Keys.ToList().ForEach(key =>
            {
                expectedFieldsDictionary[key].Should().Be(fieldsDictionary[key]);
            });
        }

        [Trait("Category", "SharePointCore - Extensions")]
        [Fact(DisplayName = "Gets all SharePoint references presents in object")]
        public void SharePointExtensions_GetReferences_AllReferencesFromObject()
        {
            // Arrange
            var item = new SharePointListItem
            {
                LookupField = new SharePointAggregatingListItem
                {
                    Id = 1
                }
            };

            var expectedReferencesDictionary = new Dictionary<string, object>();
            expectedReferencesDictionary.Add("LookupField", new SharePointAggregatingListItem { Id = 1 });

            // Act
            var referencesDictionary = item.GetReferences();

            // Assert
            expectedReferencesDictionary.Count.Should().Be(referencesDictionary.Count);

            referencesDictionary.Keys.ToList().ForEach(key =>
            {
                ((SharePointAggregatingListItem)expectedReferencesDictionary[key])
                    .Id.Should().Be(((SharePointAggregatingListItem)referencesDictionary[key]).Id);
            });
        }

        [Trait("Category", "SharePointCore - Extensions")]
        [Fact(DisplayName = "Gets all users associated from SharePoint object")]
        public void SharePointExtensions_GetUserFields_AllUserFieldsFromObject()
        {
            // Arrange
            var item = new SharePointListItem
            {
                PersonOrGroupField = "Username"
            };

            var expectedUsersDictionary = new Dictionary<string, object>();
            expectedUsersDictionary.Add("PersonOrGroupField", "Username");

            // Act
            var usersDictionary = item.GetUserFields();

            // Assert
            usersDictionary.Keys.ToList().ForEach(key =>
            {
                expectedUsersDictionary[key].Should().Be(usersDictionary[key]);
            });
        }

        [Trait("Category", "SharePointCore - Extensions")]
        [Fact(DisplayName = "Configures all aggregations fields from SharePoint object")]
        public void SharePointExtensions_ConfigureReferences_ConfiguresAggregatingFields()
        {
            // Arrange
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

            // Act
            itemListTracking.Fields["LookupField"].Should().BeNull();
            itemListTracking.ConfigureReferences(tracking);

            // Assert
            itemListTracking.Fields["LookupField"].Should().Be(aggregatingListItem.Id);
        }

        [Trait("Category", "SharePointCore - Extensions")]
        [Fact(DisplayName = "Configures all aggregations from original fields from SharePoint object")]
        public void SharePointExtensions_ConfigureReferences_ConfiguresAggregatingFromOriginalFields()
        {
            // Arrange
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

            // Act
            itemListTracking.OriginalFields["LookupField"].Should().BeNull();
            itemListTracking.ConfigureReferences(tracking, true);

            // Assert
            itemListTracking.OriginalFields["LookupField"].Should().Be(aggregatingListItem.Id);
        }
    }
}
