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

using Polimorfismo.SharePoint.Transaction;
using Polimorfismo.SharePoint.Transaction.Utils;

namespace Polimorfismo.SharePointOnline.Transaction.Tests
{
    /// <summary>
    /// Represents an item in an aggregating list.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-10 07:41:58 PM</Date>
    internal class SharePointAggregatingListItem : ISharePointItem
    {
        public string ListName => "AggregatingList";

        public int Id { get; set; }

        [SharePointField(SharePointConstants.FieldNameTitle)]
        public string TitleField { get; set; }
    }
}
