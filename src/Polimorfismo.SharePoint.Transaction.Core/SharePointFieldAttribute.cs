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
    /// Decorates a property to indicate its relation to a field present in the custom list.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 09:05:38 PM</Date>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public sealed class SharePointFieldAttribute : Attribute
    {
        #region Properties

        /// <summary>
        /// Type associated with the attribute for automatic casting.
        /// </summary>
        public Type Type { get; set; }

        /// <summary>
        /// Field Name in SharePoint.
        /// </summary>
        public string Name { get;}

        /// <summary>
        /// Indicates that the value present in this property will be used to fill 
        /// in another field that represents the association with another item present in SharePoint.
        /// </summary>
        public bool IsReference { get; set; }

        /// <summary>
        /// If the type is Lookup it can be returning both the ID or the value that represents it.
        /// </summary>
        public bool IsLookupValue { get; set; }

        #endregion

        #region Constructors / Finalizers

        public SharePointFieldAttribute(string name)
        {
            Name = name;
        }

        #endregion

        #region Methods

        internal bool IsIgnoreToInsertOrUpdate => IsReference || IsLookupValue;

        #endregion
    }
}
