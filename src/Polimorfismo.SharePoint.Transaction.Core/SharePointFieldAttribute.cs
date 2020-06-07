using System;

namespace Polimorfismo.SharePoint.Transaction.Core
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
