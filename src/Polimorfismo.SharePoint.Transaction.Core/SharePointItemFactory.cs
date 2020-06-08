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

namespace Polimorfismo.SharePoint.Transaction
{
    /// <summary>
    /// Creates new instances of SharePointItem.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-07 08:46:44 PM</Date>
    internal static class SharePointItemFactory
    {
        public static TSharePointItem Create<TSharePointItem>()
            where TSharePointItem : ISharePointItem, new()
        {
            return Activator.CreateInstance<TSharePointItem>();
        }
    }
}
