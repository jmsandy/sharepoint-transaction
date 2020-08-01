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

using System;
using Polimorfismo.SharePoint.Transaction.Utils;

namespace Polimorfismo.SharePoint.Transaction.Commons.Tests
{
    /// <summary>
    /// Represents a list item in SharePoint.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-07 09:23:35 PM</Date>
    public class SharePointListItem : ISharePointItem
    {
        public string ListName => "MyCustomList";

        public int Id { get; set; }

        [SharePointField("TextArea")]
        public string TextArea { get; set; }

        [SharePointField("TextField")]
        public string TextField { get; set; }

        [SharePointField("ChoiceField")]
        public bool ChoiceField { get; set; }

        [SharePointField("IntegerField", Type = typeof(int))]
        public int IntegerField { get; set; }

        [SharePointField("LinkField")]
        public string LinkField { get; set; }

        [SharePointField("ImageField")]
        public string ImageField { get; set; }

        [SharePointField(SharePointConstants.FieldNameTitle)]
        public string TitleField { get; set; }

        [SharePointField("OptionField")]
        public string OptionField { get; set; }

        [SharePointField("DateField")]
        public DateTime? DateField { get; set; }

        [SharePointField("DecimalField", Type = typeof(decimal))]
        public decimal DecimalField { get; set; }

        [SharePointField("CurrencyField", Type = typeof(decimal))]
        public decimal CurrencyField { get; set; }

        [SharePointField("PersonOrGroupField", IsUserValue = true)]
        public string PersonOrGroupField { get; set; }

        [SharePointField(SharePointConstants.FieldNameCreated)]
        public DateTime Created { get; set; }

        [SharePointField(SharePointConstants.FieldNameModified)]
        public DateTime Modified { get; set; }

        [SharePointField("LookupField", IsLookupValue = false)]
        public int? LookupFieldId { get; set; }

        [SharePointField("LookupField", IsReference = true)]
        public SharePointAggregatingListItem LookupField { get; set; }
    }
}
