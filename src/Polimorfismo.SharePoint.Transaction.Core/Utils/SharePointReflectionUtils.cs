using System.Linq;
using System.Reflection;
using System.Collections.Generic;

namespace Polimorfismo.SharePoint.Transaction.Core.Utils
{
    /// <summary>
    /// Extension to manipulate library metadata.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 09:07:56 PM</Date>
    internal static class SharePointReflectionUtils
    {
        #region Methods

        public static Dictionary<string, object> GetSharePointDictionaryValues<TSharePointItem>(TSharePointItem item) where TSharePointItem : ISharePointItem
        {
            var dictionary = new Dictionary<string, object>();

            foreach (var property in item.GetType()
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.GetCustomAttributes<SharePointFieldAttribute>().Any(a => !a.IsIgnoreToInsertOrUpdate)).ToList())
            {
                dictionary.Add(property.GetCustomAttributes<SharePointFieldAttribute>().First().Name, property.GetValue(item));
            }

            return dictionary;
        }

        public static Dictionary<string, object> GetSharePointRefencesDictionaryValues<TSharePointItem>(TSharePointItem item) where TSharePointItem : ISharePointItem
        {
            var dictionary = new Dictionary<string, object>();

            foreach (var property in item.GetType()
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.GetCustomAttributes<SharePointFieldAttribute>().Any(a => a.IsReference)).ToList())
            {
                dictionary.Add(property.GetCustomAttributes<SharePointFieldAttribute>().First().Name, property.GetValue(item));
            }

            return dictionary;
        }

        #endregion
    }
}
