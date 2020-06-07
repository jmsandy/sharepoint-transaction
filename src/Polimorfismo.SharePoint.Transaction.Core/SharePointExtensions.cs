using System.Collections.Generic;
using Polimorfismo.SharePoint.Transaction.Core.Utils;

namespace Polimorfismo.SharePoint.Transaction.Core
{
    /// <summary>
    /// Extension class to assist operations performed in SharePoint.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 09:36:16 PM</Date>
    internal static class SharePointExtensions
    {
        #region Extensions - ISharePointItem

        public static Dictionary<string, object> ToDictionary(this ISharePointItem item)
        {
            return SharePointReflectionUtils.GetSharePointDictionaryValues(item);
        }

        public static Dictionary<string, object> GetReferences(this ISharePointItem item)
        {
            return SharePointReflectionUtils.GetSharePointRefencesDictionaryValues(item);
        }

        public static IReadOnlyDictionary<string, object> ConfigureReferences(this SharePointItemTracking itemTracking, SharePointListItemTracking listTracking)
        {
            var fields = itemTracking.Fields.ToDictionary();

            foreach (var reference in itemTracking.Item.GetReferences())
            {
                if (fields.ContainsKey(reference.Key))
                {
                    var value = listTracking.Get(reference.Value as ISharePointItem);
                    if (value != null)
                    {
                        itemTracking.Fields[reference.Key] = value.Id;
                    }
                }
            }

            return fields;
        }

        #endregion
    }
}
