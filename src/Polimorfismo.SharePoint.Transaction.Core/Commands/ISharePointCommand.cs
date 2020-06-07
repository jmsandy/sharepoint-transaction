using System;
using System.Threading.Tasks;

namespace Polimorfismo.SharePoint.Transaction.Core.Commands
{
    /// <summary>
    /// Base interface for commands associated with operations performed over on custom lists.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 08:18:20 PM</Date>
    internal interface ISharePointCommand : IDisposable
    {
        Guid Id { get; }

        SharePointClientBase SharePointClient { get; }

        SharePointItemTracking SharePointItemTracking { get; }

        Task Prepare();

        Task Execute();

        Task Undo();
    }
}
