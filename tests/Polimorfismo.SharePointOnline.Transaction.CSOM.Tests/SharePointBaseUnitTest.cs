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
using System.Security;
using Microsoft.SharePoint.Client;
using Polimorfismo.Microsoft.SharePoint.Transaction;

namespace Polimorfismo.SharePointOnline.Transaction.Tests
{
    /// <summary>
    /// Common operations for the tests performed.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-07 09:41:03 PM</Date>
    public abstract class SharePointBaseUnitTest : IDisposable
    {
        #region Constants

        protected const string UserPwd = "";

        protected const string Username = "";

        protected const string WebFullUrl = "";

        #endregion

        #region Fields

        protected readonly SharePointClient _sharePointClient;

        private readonly SecureString _password = new SecureString();

        #endregion

        #region Constructors / Finalizers

        protected SharePointBaseUnitTest()
        {
            foreach (char c in UserPwd.ToCharArray()) _password.AppendChar(c);
            _sharePointClient = new SharePointClient(WebFullUrl, new SharePointOnlineCredentials(Username, _password));
        }

        ~SharePointBaseUnitTest() => Dispose(false);

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
                _sharePointClient?.Dispose();
            }
        }

        #endregion
    }
}
