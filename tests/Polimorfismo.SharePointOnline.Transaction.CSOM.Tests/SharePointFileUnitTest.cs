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

using Xunit;
using System;
using System.IO;
using System.Threading.Tasks;

namespace Polimorfismo.SharePointOnline.Transaction.Tests
{
    /// <summary>
    /// Testing operations to insert, update, remove and retrieve items in a SharePoint document library.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-14 08:44:00 AM</Date>
    public class SharePointFileUnitTest : SharePointBaseUnitTest
    {
        [Fact]
        public async Task SharePoint_AddFile_Root_Folder_Success_Test()
        {
            var document = new SharePointFile
            {
                FileName = $"{Guid.NewGuid()}.txt",
                Folder = "", // null or empty to add on the root folder.
                InputStream = File.OpenRead("SharePointFile.txt"),
                Description = "Document description. Others fields can be added on the document."
            };

            _sharePointClient.AddFile(document);

            await _sharePointClient.SaveChanges();
        }

        [Fact]
        public async Task SharePoint_AddFile_Multiples_Folders_Success_Test()
        {
            var document = new SharePointFile
            {
                FileName = $"{Guid.NewGuid()}.txt",
                Folder = "MyFolder/SubFolder/FileFolder",
                InputStream = File.OpenRead("SharePointFile.txt"),
                Description = "Document description. Others fields can be added on the document."
            };

            _sharePointClient.AddFile(document);

            await _sharePointClient.SaveChanges();
        }
    }
}
