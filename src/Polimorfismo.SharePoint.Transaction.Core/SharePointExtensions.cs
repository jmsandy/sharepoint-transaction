﻿// Copyright 2020 Polimorfismo - José Mauro da Silva Sandy
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

using System.Threading.Tasks;
using System.Collections.Generic;
using Polimorfismo.SharePoint.Transaction.Utils;

namespace Polimorfismo.SharePoint.Transaction
{
    /// <summary>
    /// Extension class to assist operations performed in SharePoint.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 09:36:16 PM</Date>
    internal static class SharePointExtensions
    {
        #region Extensions - ISharePointItem

        public static Dictionary<string, object> ToDictionary(this ISharePointMetadata item)
        {
            return SharePointReflectionUtils.GetSharePointFieldsDictionary(item);
        }

        public static Dictionary<string, object> GetReferences(this ISharePointMetadata item)
        {
            return SharePointReflectionUtils.GetSharePointReferencesDictionary(item);
        }

        public static Dictionary<string, object> GetUserFields(this ISharePointMetadata item)
        {
            return SharePointReflectionUtils.GetSharePointUsersDictionary(item);
        }

        public static async Task ConfigureUserFieldsAsync(this SharePointItemTracking itemTracking, SharePointClientBase sharePointClient)
        {
            foreach (var userField in itemTracking.Item.GetUserFields())
            {
                if (!string.IsNullOrWhiteSpace(userField.Value as string))
                {
                    itemTracking.Fields[userField.Key] =
                        (await sharePointClient.GetUserByLoginAsync(userField.Value as string)).Id;
                }
            }
        }

        public static IReadOnlyDictionary<string, object> ConfigureReferences(this SharePointItemTracking itemTracking, 
            SharePointListItemTracking listTracking, bool isOriginalFields = false)
        {
            var references = (isOriginalFields ? itemTracking.OriginalItem : itemTracking.Item).GetReferences();
            var fields = isOriginalFields ? itemTracking.OriginalFields.ToDictionary() : itemTracking.Fields.ToDictionary();

            foreach (var reference in references)
            {
                if (fields.ContainsKey(reference.Key))
                {
                    var value = listTracking.Get(reference.Value as ISharePointItem, isOriginalFields);
                    if (value != null)
                    {
                        if (isOriginalFields)
                        {
                            itemTracking.OriginalFields[reference.Key] = value.OriginalItem.Id;
                        }
                        else
                        {
                            itemTracking.Fields[reference.Key] = value.Id;
                        }
                    }
                }
            }

            return fields;
        }

        #endregion
    }
}
