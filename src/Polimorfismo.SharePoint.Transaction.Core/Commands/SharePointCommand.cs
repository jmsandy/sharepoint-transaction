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

using System;
using System.Threading.Tasks;

namespace Polimorfismo.SharePoint.Transaction.Commands
{
    /// <summary>
    /// Base implementation of a command to be executed to communicate with custom lists.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-05-24 08:22:42 PM</Date>
    internal abstract class SharePointCommand<TSharePointItem> : ISharePointCommand where TSharePointItem : ISharePointItem
    {
        #region Properties

        public Guid Id { get; }

        public SharePointClientBase SharePointClient { get; }

        public SharePointItemTracking SharePointItemTracking { get; }

        #endregion

        #region Constructors / Finalizers

        protected SharePointCommand(SharePointClientBase sharePointClient, SharePointItemTracking itemTracking)
        {
            Id = Guid.NewGuid();
            SharePointClient = sharePointClient;
            SharePointItemTracking = itemTracking;
        }

        ~SharePointCommand() => Dispose(false);

        #endregion

        #region ISharePointCommand - Members

        public abstract Task Prepare();

        public abstract Task Execute();

        public abstract Task Undo();

        #endregion

        #region IDisposable - Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public virtual void Dispose(bool disposing)
        {
            if (disposing)
            {   
                SharePointClient?.Dispose();
            }
        }

        #endregion
    }
}
