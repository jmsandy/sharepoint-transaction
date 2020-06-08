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
using Polimorfismo.SharePoint.Transaction;
using Polimorfismo.SharePoint.Transaction.Utils;

namespace Polimorfismo.SharePointOnline.Transaction.Tests
{
    /// <summary>
    /// Represents a list item in SharePoint.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-07 09:23:35 PM</Date>
    internal class SharePointListItem : ISharePointItem
    {
        public string ListName => "MyCustomList";

        public int Id { get; set; }

        [SharePointField("Title")]
        public string TitleField { get; set; }

        [SharePointField("TextField")]
        public string TextField { get; set; }

        [SharePointField("TextArea")]
        public string TextArea { get; set; }

        [SharePointField("OptionField")]
        public string OptionField { get; set; }

        [SharePointField("IntegerField",  Type = typeof(int))]
        public int IntegerField { get; set; }

        [SharePointField("DecimalField", Type = typeof(decimal))]
        public decimal DecimalField { get; set; }

        [SharePointField(SharePointConstants.FieldNameCreated)]
        public DateTime Created { get; set; }

        [SharePointField(SharePointConstants.FieldNameModified)]
        public DateTime Modified { get; set; }
    }
}
