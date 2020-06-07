using Polimorfismo.SharePoint.Transaction.Core.Utils;

namespace Polimorfismo.SharePoint.Transaction.Core
{
    /// <summary>
    /// Interface representing an item in the custom list on the SharePoint. 
    /// It can be having a representation for each operation to be performed.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 08:46:57 PM</Date>
    public interface ISharePointItem
    {
        [SharePointField(SharePointConstants.FieldNameId)]
        int Id { get; set; }

        string ListName { get; }
    }
}
