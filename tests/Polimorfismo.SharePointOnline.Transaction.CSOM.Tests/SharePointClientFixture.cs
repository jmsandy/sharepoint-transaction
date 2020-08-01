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

using Bogus;
using Xunit;
using System;
using System.Security;
using Microsoft.SharePoint.Client;
using Polimorfismo.SharePoint.Transaction;
using Polimorfismo.SharePoint.Transaction.Commons.Tests;

namespace Polimorfismo.SharePointOnline.Transaction.Tests
{
    /// <summary>
    /// Collection with shared resources to SharePoint tests.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-07-29 06:46:51 PM</Date>
    [CollectionDefinition(nameof(SharePointClientCollection))]
    public class SharePointClientCollection : ICollectionFixture<SharePointClientFixture>
    {
    }

    /// <summary>
    /// Auxiliary class with shared resources to SharePoint tests.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-07-29 06:42:34 PM</Date>
    public class SharePointClientFixture : IDisposable
    {
        #region Constants

        protected const int UserId = 0;

        protected const string UserPwd = "";

        protected const string Username = "";

        protected const string WebFullUrl = "";

        #endregion

        #region Fields

        protected SharePointClient _sharePointClient;

        private readonly SecureString _password = new SecureString();

        #endregion

        #region Properties

        protected bool Disposed { get; private set; }

        public SharePointClient SharePointClient
        {
            get
            {
                if (_sharePointClient == null)
                {
                    foreach (char c in UserPwd.ToCharArray()) _password.AppendChar(c);
                    _sharePointClient = new SharePointClient(WebFullUrl, new SharePointOnlineCredentials(Username, _password));
                }

                return _sharePointClient;
            }
        }

        #endregion

        #region Constructors / Finalizers

        ~SharePointClientFixture() => Dispose(false);

        #endregion

        #region Methods

        public SharePointListItem GenerateSharePointListItem(bool validItem = true)
        {
            return new Faker<SharePointListItem>()
                .RuleFor(item => item.OptionField, (f, c) => "Option 2")
                .RuleFor(item => item.TextArea, (f, c) => f.Lorem.Lines(2))
                .RuleFor(item => item.TextField, (f, c) => f.Lorem.Lines(1))
                .RuleFor(item => item.LinkField, (f, c) => f.Internet.Url())
                .RuleFor(item => item.IntegerField, (f, c) => f.Random.Int())
                .RuleFor(item => item.ChoiceField, (f, c) => f.Random.Bool())
                .RuleFor(item => item.PersonOrGroupField, (f, c) => Username)
                .RuleFor(item => item.TitleField, (f, c) => f.Lorem.Sentence(2))
                .RuleFor(item => item.DateField, (f, c) => f.Date.Recent().Date)
                .RuleFor(item => item.DecimalField, (f, c) => f.Random.Decimal())
                .RuleFor(item => item.CurrencyField, (f, c) => f.Random.Decimal())
                .RuleFor(item => item.ImageField, (f, c) => f.Image.LoremFlickrUrl())
                .RuleFor(item => item.TitleField, (f, c) => validItem ? f.Lorem.Sentence(2) : f.Lorem.Letter(300))
                .Generate();
        }

        public SharePointAggregatingListItem GenerateSharePointAggregatingListItem()
        {
            return new Faker<SharePointAggregatingListItem>()
                .RuleFor(item => item.TitleField, (f, c) => f.Lorem.Sentence(2))
                .Generate();
        }

        public SharePointFile GenerateSharePointFile(string folder = "", bool validFile = true)
        {
            return new Faker<SharePointFile>()
                .RuleFor(item => item.Folder, (f, c) => folder)
                .RuleFor(item => item.FileName, (f, c) => f.System.FileName("txt"))
                .RuleFor(item => item.InputStream, (f, c) => System.IO.File.OpenRead("SharePointFile.txt"))
                .RuleFor(item => item.Description, (f, c) => validFile ? f.Lorem.Sentence(3) : f.Lorem.Letter(300))
                .Generate();
        }

        public string GetAbsolutePath(string relativePath) => $"{WebFullUrl}/{relativePath}";

        public string GetRelativePath(string relativePath) => $"/sites/develop/{relativePath}";

        #endregion

        #region IDisposable - Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (Disposed) return;

            if (disposing)
            {
                _sharePointClient?.Dispose();
            }

            Disposed = true;
        }

        #endregion
    }
}
