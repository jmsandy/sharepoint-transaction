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
using Polimorfismo.SharePoint.Transaction.Utils;
using Polimorfismo.SharePoint.Transaction.Commons.Tests;

namespace Polimorfismo.SharePoint.Transaction.Core.Tests
{
    /// <summary>
    /// Tests related to the reflection process.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-15 08:06:46 PM</Date>
    public class SharePointReflectionTests
    {
        [Trait("Category", "SharePointCore - Reflection")]
        [Fact(DisplayName = "Retrieves all fields from SharePoint objects")]
        public void SharePointReflectionUtils_GetSharePointFieldsDictionary_FieldsRetrievedWithSuccess()
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
            var fieldsDictionary = SharePointReflectionUtils.GetSharePointFieldsDictionary(item);

            // Assert
            expectedFieldsDictionary.Count.Should().Be(fieldsDictionary.Count);

            fieldsDictionary.Keys.ToList().ForEach(key =>
            {
                expectedFieldsDictionary[key].Should().Be(fieldsDictionary[key]);
            });
        }

        [Trait("Category", "SharePointCore - Reflection")]
        [Fact(DisplayName = "Retrieves references from SharePoint objects")]
        public void SharePointReflectionUtils_GetSharePointReferencesDictionary_ReferencesRetrievedWithSuccess()
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
            var referencesDictionary = SharePointReflectionUtils.GetSharePointReferencesDictionary(item);

            // Assert
            expectedReferencesDictionary.Count.Should().Be(referencesDictionary.Count);

            referencesDictionary.Keys.ToList().ForEach(key =>
            {
                ((SharePointAggregatingListItem)expectedReferencesDictionary[key])
                    .Id.Should().Be(((SharePointAggregatingListItem)referencesDictionary[key]).Id);
            });
        }

        [Trait("Category", "SharePointCore - Reflection")]
        [Fact(DisplayName = "Retrieves users from SharePoint objects")]
        public void SharePointReflectionUtils_GetSharePointUsersDictionary_UsersRetrievedWithSuccess()
        {
            // Arrange
            var item = new SharePointListItem
            {
                PersonOrGroupField = "Username"
            };

            var expectedUsersDictionary = new Dictionary<string, object>();
            expectedUsersDictionary.Add("PersonOrGroupField", "Username");

            // Act
            var usersDictionary = SharePointReflectionUtils.GetSharePointUsersDictionary(item);

            // Assert
            usersDictionary.Keys.ToList().ForEach(key =>
            {
                expectedUsersDictionary[key].Should().Be(usersDictionary[key]);
            });
        }
    }
}
