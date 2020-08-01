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

using Xunit;
using System;
using System.IO;
using CamlexNET;
using System.Linq;
using FluentAssertions;
using System.Threading.Tasks;
using Polimorfismo.SharePoint.Transaction;
using Polimorfismo.SharePoint.Transaction.Utils;
using Polimorfismo.SharePoint.Transaction.Commons.Tests;
using System.Xml.XPath;

namespace Polimorfismo.SharePointOnline.Transaction.Tests
{
    /// <summary>
    /// Testing operations to insert, update, remove and retrieve items in a SharePoint custom list.
    /// </summary>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-07 09:36:53 PM</Date>
    [Collection(nameof(SharePointClientCollection))]
    public class SharePointClientTests
    {
        private readonly SharePointClientFixture _sharePointClientFixture;

        public SharePointClientTests(SharePointClientFixture sharePointClientFixture)
        {
            _sharePointClientFixture = sharePointClientFixture;
        }

        [Trait("Category", "SharePointOnline - User")]
        [Theory(DisplayName = "Exception when getting the user when the login is null or empty")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("    ")]
        public void SharePointClient_GetUserByLogin_LoginNullOrEmpty(string login)
        {
            // Arrange & Act & Asert
            var exception = Assert.Throws<ArgumentNullException>(() 
                => _sharePointClientFixture.SharePointClient.GetUserByLogin(login));

            exception.Message.Should().Contain("login");
        }

        [Trait("Category", "SharePointOnline - Items")]
        [Fact(DisplayName = "Adding single item in SharePoint list")]
        public async Task SharePointClient_AddItem_SingleItem()
        {
            // Arrange
            var item = _sharePointClientFixture.GenerateSharePointListItem();

            // Act
            _sharePointClientFixture.SharePointClient.AddItem(item);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var expectedItem = await _sharePointClientFixture.SharePointClient
                .GetItemByIdAsync<SharePointListItem>(item.Id);

            // Assert
            expectedItem.Should().NotBeNull();
            item.Id.Should().Be(expectedItem.Id);
            item.TextArea.Should().Be(expectedItem.TextArea);
            item.LinkField.Should().Be(expectedItem.LinkField);
            item.TextField.Should().Be(expectedItem.TextField);
            item.TitleField.Should().Be(expectedItem.TitleField);
            item.ImageField.Should().Be(expectedItem.ImageField);
            item.ChoiceField.Should().Be(expectedItem.ChoiceField);
            item.OptionField.Should().Be(expectedItem.OptionField);
            item.DecimalField.Should().Be(expectedItem.DecimalField);
            item.IntegerField.Should().Be(expectedItem.IntegerField);
            item.CurrencyField.Should().Be(expectedItem.CurrencyField);
            item.DateField.Should().Be(expectedItem.DateField.Value.Date);
        }

        [Trait("Category", "SharePointOnline - Items")]
        [Fact(DisplayName = "Adding item in SharePoint list with object reference")]
        public async Task SharePointClient_AddItem_WithAssociationByObjectReference()
        {
            // Arrange
            var aggregatingListItem = _sharePointClientFixture.GenerateSharePointAggregatingListItem();

            var listItem = _sharePointClientFixture.GenerateSharePointListItem();
            listItem.LookupField = aggregatingListItem;

            // Act
            _sharePointClientFixture.SharePointClient.AddItem(listItem);
            _sharePointClientFixture.SharePointClient.AddItem(aggregatingListItem);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var expectedListItem = await _sharePointClientFixture.SharePointClient.GetItemByIdAsync<SharePointListItem>(listItem.Id);

            // Assert
            expectedListItem.TitleField.Should().Be(listItem.TitleField);
            expectedListItem.LookupFieldId.Should().Be(aggregatingListItem.Id);
        }

        [Trait("Category", "SharePointOnline - Items")]
        [Fact(DisplayName = "Adding an item associated with SharePoint with a known id")]
        public async Task SharePointClient_AddItem_WithAssociationById()
        {
            // Arrange
            var aggregatingListItem = _sharePointClientFixture.GenerateSharePointAggregatingListItem();

            // Act
            _sharePointClientFixture.SharePointClient.AddItem(aggregatingListItem);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var listItem = _sharePointClientFixture.GenerateSharePointListItem();
            var listItem2 = _sharePointClientFixture.GenerateSharePointListItem();

            listItem.LookupFieldId = aggregatingListItem.Id;
            listItem2.LookupFieldId = aggregatingListItem.Id;

            _sharePointClientFixture.SharePointClient.AddItem(listItem);
            _sharePointClientFixture.SharePointClient.AddItem(listItem2);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var items = await _sharePointClientFixture.SharePointClient.GetItemsAsync<SharePointListItem>(
                Camlex.Query().Where(i => i["LookupField"] == (DataTypes.LookupId)aggregatingListItem.Id.ToString()).ToCamlQuery());

            // Assert
            items.Count.Should().Be(2);
        }

        [Trait("Category", "SharePointOnline - Items")]
        [Fact(DisplayName = "Rollback when inserting a second invalid item")]
        public async Task SharePointClient_AddItem_Rollback()
        {
            // Arrange
            var item = _sharePointClientFixture.GenerateSharePointListItem();
            var invalidItem = _sharePointClientFixture.GenerateSharePointListItem(false);

            // Act
            _sharePointClientFixture.SharePointClient.AddItem(item);
            _sharePointClientFixture.SharePointClient.AddItem(invalidItem);

            var exception = await Assert.ThrowsAsync<SharePointException>(()
               => _sharePointClientFixture.SharePointClient.SaveChangesAsync());
            var expectedItem = await _sharePointClientFixture.SharePointClient
                .GetItemByIdAsync<SharePointListItem>(item.Id);

            // Assert
            expectedItem.Should().BeNull();
            exception.ErrorCode.Should().Be(SharePointErrorCode.SaveChanges);
        }

        [Trait("Category", "SharePointOnline - Items")]
        [Fact(DisplayName = "Updating item in SharePoint list")]
        public async Task SharePointClient_UpdateItem_Success()
        {
            // Arrange
            var item = _sharePointClientFixture.GenerateSharePointListItem();

            // Act
            _sharePointClientFixture.SharePointClient.AddItem(item);
            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var updatedItem = await _sharePointClientFixture.SharePointClient.GetItemByIdAsync<SharePointListItem>(item.Id);

            var updatedTitle = $"Updated - {item.TitleField}";
            updatedItem.TitleField = updatedTitle;

            _sharePointClientFixture.SharePointClient.UpdateItem(updatedItem);
            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var expectedItem = await _sharePointClientFixture.SharePointClient.GetItemByIdAsync<SharePointListItem>(updatedItem.Id);
            updatedItem.TitleField = updatedTitle;

            // Assert
            updatedTitle.Should().Be(expectedItem.TitleField);
        }

        [Trait("Category", "SharePointOnline - Items")]
        [Fact(DisplayName = "Rollback when updating a second invalid item")]
        public async Task SharePointClient_UpdateItem_Rollback()
        {
            // Arrange
            var item = _sharePointClientFixture.GenerateSharePointListItem();
            var invalidItem = _sharePointClientFixture.GenerateSharePointListItem(false);
            var expectedTitle = item.TitleField;

            // Act
            _sharePointClientFixture.SharePointClient.AddItem(item);
            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            item.TitleField = "Updated title";

            _sharePointClientFixture.SharePointClient.UpdateItem(item);
            _sharePointClientFixture.SharePointClient.AddItem(invalidItem);

            var exception = await Assert.ThrowsAsync<SharePointException>(()
               => _sharePointClientFixture.SharePointClient.SaveChangesAsync());

            var expectedItem = await _sharePointClientFixture.SharePointClient
                .GetItemByIdAsync<SharePointListItem>(item.Id);

            var expectedInvalidItem = await _sharePointClientFixture.SharePointClient
                .GetItemByIdAsync<SharePointListItem>(invalidItem.Id);

            // Assert
            expectedInvalidItem.Should().BeNull();
            expectedItem.TitleField.Should().Be(expectedTitle);
            exception.ErrorCode.Should().Be(SharePointErrorCode.SaveChanges);
        }

        [Trait("Category", "SharePointOnline - Items")]
        [Fact(DisplayName = "Delete item in SharePoint list")]
        public async Task SharePointClient_DeleteItem_Success()
        {
            // Arrange
            var item = _sharePointClientFixture.GenerateSharePointListItem();

            // Act
            _sharePointClientFixture.SharePointClient.AddItem(item);
            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var deletedItem = await _sharePointClientFixture.SharePointClient.GetItemByIdAsync<SharePointListItem>(item.Id);

            _sharePointClientFixture.SharePointClient.DeleteItem(deletedItem);
            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var expectedItem = await _sharePointClientFixture.SharePointClient.GetItemByIdAsync<SharePointListItem>(item.Id);

            // Assert
            expectedItem.Should().BeNull();
        }

        [Trait("Category", "SharePointOnline - Items")]
        [Fact(DisplayName = "Rollback when deleting a second invalid item")]
        public async Task SharePointClient_DeleteItem_Rollback()
        {
            // Arrange
            var item = _sharePointClientFixture.GenerateSharePointListItem();
            var invalidItem = _sharePointClientFixture.GenerateSharePointListItem(false);

            // Act
            _sharePointClientFixture.SharePointClient.AddItem(item);
            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            _sharePointClientFixture.SharePointClient.DeleteItem(item);
            _sharePointClientFixture.SharePointClient.AddItem(invalidItem);

            var exception = await Assert.ThrowsAsync<SharePointException>(()
               => _sharePointClientFixture.SharePointClient.SaveChangesAsync());

            var expectedItem = (await _sharePointClientFixture.SharePointClient.GetItemsAsync<SharePointListItem>(
                Camlex.Query().Where(i => (string)i[SharePointConstants.FieldNameTitle] == item.TitleField).ToCamlQuery())).FirstOrDefault();

            var expectedInvalidItem = await _sharePointClientFixture.SharePointClient
                .GetItemByIdAsync<SharePointListItem>(invalidItem.Id);

            // Assert
            expectedInvalidItem.Should().BeNull();
            expectedItem.TitleField.Should().Be(expectedItem.TitleField);
            exception.ErrorCode.Should().Be(SharePointErrorCode.SaveChanges);
        }

        [Trait("Category", "SharePointOnline - Items")]
        [Fact(DisplayName = "[Async] Gets SharePoint item by id")]
        public async Task SharePointClient_GetItemByIdAsync_Success()
        {
            // Arrange
            var item = _sharePointClientFixture.GenerateSharePointListItem();

            // Act
            _sharePointClientFixture.SharePointClient.AddItem(item);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var expectedItem = await _sharePointClientFixture.SharePointClient.GetItemByIdAsync<SharePointListItem>(item.Id);

            // Assert
            expectedItem.Should().NotBeNull();
            item.Id.Should().Be(expectedItem.Id);
            item.TextArea.Should().Be(expectedItem.TextArea);
            item.LinkField.Should().Be(expectedItem.LinkField);
            item.TextField.Should().Be(expectedItem.TextField);
            item.TitleField.Should().Be(expectedItem.TitleField);
            item.ImageField.Should().Be(expectedItem.ImageField);
            item.ChoiceField.Should().Be(expectedItem.ChoiceField);
            item.OptionField.Should().Be(expectedItem.OptionField);
            item.DecimalField.Should().Be(expectedItem.DecimalField);
            item.IntegerField.Should().Be(expectedItem.IntegerField);
            item.CurrencyField.Should().Be(expectedItem.CurrencyField);
            item.DateField.Should().Be(expectedItem.DateField.Value.Date);
        }

        [Trait("Category", "SharePointOnline - Items")]
        [Fact(DisplayName = "[Sync] Gets SharePoint item by id")]
        public void SharePointClient_GetItemById_Success()
        {
            // Arrange
            var item = _sharePointClientFixture.GenerateSharePointListItem();

            // Act
            _sharePointClientFixture.SharePointClient.AddItem(item);

            _sharePointClientFixture.SharePointClient.SaveChanges();

            var expectedItem = _sharePointClientFixture.SharePointClient.GetItemById<SharePointListItem>(item.Id);

            // Assert
            expectedItem.Should().NotBeNull();
            item.Id.Should().Be(expectedItem.Id);
            item.TextArea.Should().Be(expectedItem.TextArea);
            item.LinkField.Should().Be(expectedItem.LinkField);
            item.TextField.Should().Be(expectedItem.TextField);
            item.TitleField.Should().Be(expectedItem.TitleField);
            item.ImageField.Should().Be(expectedItem.ImageField);
            item.ChoiceField.Should().Be(expectedItem.ChoiceField);
            item.OptionField.Should().Be(expectedItem.OptionField);
            item.DecimalField.Should().Be(expectedItem.DecimalField);
            item.IntegerField.Should().Be(expectedItem.IntegerField);
            item.CurrencyField.Should().Be(expectedItem.CurrencyField);
            item.DateField.Should().Be(expectedItem.DateField.Value.Date);
        }

        [Trait("Category", "SharePointOnline - Items")]
        [Fact(DisplayName = "[Async] Gets all SharePoint items by title")]
        public async Task SharePointClient_GetItemsAsync_ByTitle()
        {
            // Arrange
            var item = _sharePointClientFixture.GenerateSharePointListItem();

            // Act
            _sharePointClientFixture.SharePointClient.AddItem(item);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var items = await _sharePointClientFixture.SharePointClient.GetItemsAsync<SharePointListItem>(
                Camlex.Query().Where(i => (string)i[SharePointConstants.FieldNameTitle] == item.TitleField).ToCamlQuery());

            // Assert
            item.TitleField.Should().Be(items.Single().TitleField);
        }

        [Trait("Category", "SharePointOnline - Items")]
        [Fact(DisplayName = "[Sync] Gets all SharePoint items by title")]
        public void SharePointClient_GetItems_ByTitle()
        {
            // Arrange
            var item = _sharePointClientFixture.GenerateSharePointListItem();

            // Act
            _sharePointClientFixture.SharePointClient.AddItem(item);

            _sharePointClientFixture.SharePointClient.SaveChanges();

            var items = _sharePointClientFixture.SharePointClient.GetItems<SharePointListItem>(
                Camlex.Query().Where(i => (string)i[SharePointConstants.FieldNameTitle] == item.TitleField).ToCamlQuery());

            // Assert
            item.TitleField.Should().Be(items.Single().TitleField);
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "Adding a file on root folder the SharePoint document library")]
        public async Task SharePointClient_AddFile_RootFolder()
        {
            // Arrange
            var file = _sharePointClientFixture.GenerateSharePointFile();

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var expectedFile = await _sharePointClientFixture.SharePointClient
                .GetFileByIdAsync<SharePointFile>(file.Id);

            // Assert
            expectedFile.Should().NotBeNull();
            file.Id.Should().Be(expectedFile.Id);
            file.Folder.Should().Be(expectedFile.Folder);
            file.FileName.Should().Be(expectedFile.FileName);
            file.Description.Should().Be(expectedFile.Description);
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "Rollback when inserting a second invalid file inside a root folder")]
        public async Task SharePointClient_AddFile_RootFolderRollback()
        {
            // Arrange
            var file = _sharePointClientFixture.GenerateSharePointFile();
            var invalidFile = _sharePointClientFixture.GenerateSharePointFile(validFile: false);

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file);
            _sharePointClientFixture.SharePointClient.AddFile(invalidFile);

            var exception = await Assert.ThrowsAsync<SharePointException>(()
               => _sharePointClientFixture.SharePointClient.SaveChangesAsync());

            var expectedFile = await _sharePointClientFixture.SharePointClient
                .GetFileByIdAsync<SharePointFile>(file.Id);

            var expectedInvalidFile = await _sharePointClientFixture.SharePointClient
                .GetFileByIdAsync<SharePointFile>(file.Id);

            // Assert
            file.Id.Should().Be(0);
            invalidFile.Id.Should().Be(0);
            expectedFile.Should().BeNull();
            expectedInvalidFile.Should().BeNull();
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "Adding a file on custom folder the SharePoint document library")]
        public async Task SharePointClient_AddFile_CustomFolder()
        {
            // Arrange
            var file = _sharePointClientFixture.GenerateSharePointFile("MyFolder");

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var expectedFile = await _sharePointClientFixture.SharePointClient
                .GetFileByIdAsync<SharePointFile>(file.Id);

            // Assert
            expectedFile.Should().NotBeNull();
            file.Id.Should().Be(expectedFile.Id);
            file.Folder.Should().Be(expectedFile.Folder);
            file.FileName.Should().Be(expectedFile.FileName);
            file.Description.Should().Be(expectedFile.Description);
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "Rollback when inserting a second invalid file on custom folder the SharePoint document library")]
        public async Task SharePointClient_AddFile_CustomFolderRollback()
        {
            // Arrange
            var file = _sharePointClientFixture.GenerateSharePointFile("MyFolder");
            var invalidFile = _sharePointClientFixture.GenerateSharePointFile("MyFolder", false);

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file);
            _sharePointClientFixture.SharePointClient.AddFile(invalidFile);

            var exception = await Assert.ThrowsAsync<SharePointException>(()
               => _sharePointClientFixture.SharePointClient.SaveChangesAsync());

            var expectedFile = await _sharePointClientFixture.SharePointClient
                .GetFileByIdAsync<SharePointFile>(file.Id);

            var expectedInvalidFile = await _sharePointClientFixture.SharePointClient
                .GetFileByIdAsync<SharePointFile>(file.Id);

            // Assert
            file.Id.Should().Be(0);
            invalidFile.Id.Should().Be(0);
            expectedFile.Should().BeNull();
            expectedInvalidFile.Should().BeNull();
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "Adding a file on folder hierarchy the SharePoint document library")]
        public async Task SharePointClient_AddFile_FolderHierarchy()
        {
            // Arrange
            var file = _sharePointClientFixture.GenerateSharePointFile("Level0/Level1/Level2");

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var expectedFile = await _sharePointClientFixture.SharePointClient
                .GetFileByIdAsync<SharePointFile>(file.Id);

            // Assert
            expectedFile.Should().NotBeNull();
            file.Id.Should().Be(expectedFile.Id);
            file.Folder.Should().Be(expectedFile.Folder);
            file.FileName.Should().Be(expectedFile.FileName);
            file.Description.Should().Be(expectedFile.Description);
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "Rollback when inserting a second invalid file on folder hierarchy the SharePoint document library")]
        public async Task SharePointClient_AddFile_FolderHierarchyRollback()
        {
            // Arrange
            var file = _sharePointClientFixture.GenerateSharePointFile("Level0/Level1/Level2");
            var invalidFile = _sharePointClientFixture.GenerateSharePointFile("Level0/Level1/Level2", false);

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file);
            _sharePointClientFixture.SharePointClient.AddFile(invalidFile);

            var exception = await Assert.ThrowsAsync<SharePointException>(()
               => _sharePointClientFixture.SharePointClient.SaveChangesAsync());

            var expectedFile = await _sharePointClientFixture.SharePointClient
                .GetFileByIdAsync<SharePointFile>(file.Id);

            var expectedInvalidFile = await _sharePointClientFixture.SharePointClient
                .GetFileByIdAsync<SharePointFile>(file.Id);

            // Assert
            file.Id.Should().Be(0);
            invalidFile.Id.Should().Be(0);
            expectedFile.Should().BeNull();
            expectedInvalidFile.Should().BeNull();
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "Updating a file on folder the SharePoint document library")]
        public async Task SharePointClient_UpdateFile_ChangingFileContent()
        {
            // Arrange
            var file = _sharePointClientFixture.GenerateSharePointFile();

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            file.Description = "Description updated";
            file.InputStream = File.OpenRead("SharePointFileUpdated.txt");

            _sharePointClientFixture.SharePointClient.UpdateFile(file);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var remoteFile = await _sharePointClientFixture.SharePointClient.GetFileByIdAsync<SharePointFile>(file.Id);

            var remoteContentFile = new StreamReader(remoteFile.InputStream).ReadToEnd();
            var expectedContentFile = new StreamReader(File.OpenRead("SharePointFileUpdated.txt")).ReadToEnd();

            // Assert
            expectedContentFile.Should().Be(remoteContentFile);
            remoteFile.Description.Should().Be(file.Description);
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "Rollback when updating a second invalid file on folder the SharePoint document library")]
        public async Task SharePointClient_UpdateFile_ChangingFileContentRollback()
        {
            // Arrange
            var file = _sharePointClientFixture.GenerateSharePointFile();
            var invalidFile = _sharePointClientFixture.GenerateSharePointFile(validFile: false);
            var expectedDescription = file.Description;

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            file.Description = "Description updated";
            file.InputStream = File.OpenRead("SharePointFileUpdated.txt");

            _sharePointClientFixture.SharePointClient.UpdateFile(file);
            _sharePointClientFixture.SharePointClient.AddFile(invalidFile);

            var exception = await Assert.ThrowsAsync<SharePointException>(()
               => _sharePointClientFixture.SharePointClient.SaveChangesAsync());

            var expectedFile = await _sharePointClientFixture.SharePointClient
                .GetFileByIdAsync<SharePointFile>(file.Id);

            var remoteContentFile = new StreamReader(expectedFile.InputStream).ReadToEnd();
            var expectedContentFile = new StreamReader(File.OpenRead("SharePointFile.txt")).ReadToEnd();

            var expectedInvalidFile = await _sharePointClientFixture.SharePointClient
                .GetFileByIdAsync<SharePointFile>(invalidFile.Id);

            // Assert
            expectedFile.Id.Should().Be(file.Id);
            expectedInvalidFile.Should().BeNull();
            expectedContentFile.Should().Be(remoteContentFile);
            expectedFile.Description.Should().Be(expectedDescription);
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "Deleting a file on folder the SharePoint document library")]
        public async Task SharePointClient_DeleteFile_DeleteWithSuccess()
        {
            // Arrange
            var file = _sharePointClientFixture.GenerateSharePointFile();

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            _sharePointClientFixture.SharePointClient.DeleteFile(file);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var expectedFile = await _sharePointClientFixture.SharePointClient.GetFileByIdAsync<SharePointFile>(file.Id);

            // Assert
            expectedFile.Should().BeNull();
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "Rollback when deleting a second invalid file on folder the SharePoint document library")]
        public async Task SharePointClient_DeleteFile_DeleteRollback()
        {
            // Arrange
            var file = _sharePointClientFixture.GenerateSharePointFile();
            var invalidFile = _sharePointClientFixture.GenerateSharePointFile(validFile: false);

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            _sharePointClientFixture.SharePointClient.DeleteFile(file);
            _sharePointClientFixture.SharePointClient.AddFile(invalidFile);

            var exception = await Assert.ThrowsAsync<SharePointException>(()
               => _sharePointClientFixture.SharePointClient.SaveChangesAsync());

            var expectedFile = _sharePointClientFixture.SharePointClient.GetFiles<SharePointFile>(
                Camlex.Query().Where(i => (string)i[SharePointConstants.FieldNameFileLeafRef] == file.FileName).ToCamlQuery()).FirstOrDefault();

            var expectedInvalidFile = await _sharePointClientFixture.SharePointClient
                .GetFileByIdAsync<SharePointFile>(invalidFile.Id);

            // Assert
            expectedFile.Should().NotBeNull();
            expectedInvalidFile.Should().BeNull();
            expectedFile.Description.Should().Be(file.Description);
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "[Async] Gets SharePoint file by id")]
        public async Task SharePointClient_GetFileByIdAsync_Success()
        {
            // Arrange
            var file = _sharePointClientFixture.GenerateSharePointFile();

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var expectedFile = await _sharePointClientFixture.SharePointClient.GetFileByIdAsync<SharePointFile>(file.Id);

            // Assert
            expectedFile.Should().NotBeNull();
            file.Folder.Should().Be(expectedFile.Folder);
            file.FileName.Should().Be(expectedFile.FileName);
            file.Description.Should().Be(expectedFile.Description);
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "[Sync] Gets SharePoint file by id")]
        public void SharePointClient_GetFileById_Success()
        {
            // Arrange
            var file = _sharePointClientFixture.GenerateSharePointFile();

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file);

            _sharePointClientFixture.SharePointClient.SaveChanges();

            var expectedFile = _sharePointClientFixture.SharePointClient.GetFileById<SharePointFile>(file.Id);

            // Assert
            expectedFile.Should().NotBeNull();
            file.Folder.Should().Be(expectedFile.Folder);
            file.FileName.Should().Be(expectedFile.FileName);
            file.Description.Should().Be(expectedFile.Description);
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "[Async] Gets all SharePoint files by filename")]
        public async Task SharePointClient_GetFilesAsync_ByFileName()
        {
            // Arrange
            var file = _sharePointClientFixture.GenerateSharePointFile();

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var files = await _sharePointClientFixture.SharePointClient.GetFilesAsync<SharePointFile>(
                Camlex.Query().Where(i => (string)i[SharePointConstants.FieldNameFileLeafRef] == file.FileName).ToCamlQuery());

            // Assert
            file.FileName.Should().Be(files.Single().FileName);
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "[Sync] Gets all SharePoint files by filename")]
        public void SharePointClient_GetFiles_ByFileName()
        {
            // Arrange
            var file = _sharePointClientFixture.GenerateSharePointFile();

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file);

            _sharePointClientFixture.SharePointClient.SaveChanges();

            var files = _sharePointClientFixture.SharePointClient.GetFiles<SharePointFile>(
                Camlex.Query().Where(i => (string)i[SharePointConstants.FieldNameFileLeafRef] == file.FileName).ToCamlQuery());

            // Assert
            file.FileName.Should().Be(files.Single().FileName);
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "[Sync] Getting all documents information on specific folder the SharePoint document library from relative path")]
        public void SharePointClient_GetDocumentsInfo_GettingAllDocumentsFromRelativePath()
        {
            // Arrange
            var folder = Guid.NewGuid().ToString();
            var file1 = _sharePointClientFixture.GenerateSharePointFile(folder);
            var file2 = _sharePointClientFixture.GenerateSharePointFile($"{folder}/SubFolder");

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file1);
            _sharePointClientFixture.SharePointClient.AddFile(file2);

            _sharePointClientFixture.SharePointClient.SaveChanges();

            var documentInfo = _sharePointClientFixture.SharePointClient
                .GetDocumentsInfo("DocumentsList", _sharePointClientFixture.GetRelativePath($"DocumentsList/{folder}"));

            // Assert
            documentInfo.Length.Should().Be(0);
            documentInfo.Name.Should().Be(folder);
            documentInfo.IsFile.Should().BeFalse();
            documentInfo.ContentBase64.Should().BeNull();
            documentInfo.Documents.Count().Should().Be(2);
            documentInfo.Extension.Should().BeNullOrEmpty();
            documentInfo.Documents.Should().Contain(d => d.Id.Equals(file1.Id));
            documentInfo.Documents.Should().Contain(d => d.Name.Equals("SubFolder"));
            documentInfo.Documents.Should().Contain(d => d.Name.Equals(file1.FileName));

            var subFolder = documentInfo.Documents.First(d => d.Name.Equals("SubFolder"));
            subFolder.Length.Should().Be(0);
            subFolder.IsFile.Should().BeFalse();
            subFolder.ContentBase64.Should().BeNull();
            subFolder.Documents.Count().Should().Be(1);
            subFolder.Extension.Should().BeNullOrEmpty();
            subFolder.Documents.Should().Contain(d => d.Id.Equals(file2.Id));
            subFolder.Documents.Should().Contain(d => d.Name.Equals(file2.FileName));
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "[Async] Getting all documents information on specific folder the SharePoint document library from relative path")]
        public async Task SharePointClient_GetDocumentsInfoAsync_GettingAllDocumentsFromRelativePath()
        {
            // Arrange
            var folder = Guid.NewGuid().ToString();
            var file1 = _sharePointClientFixture.GenerateSharePointFile(folder);
            var file2 = _sharePointClientFixture.GenerateSharePointFile($"{folder}/SubFolder");

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file1);
            _sharePointClientFixture.SharePointClient.AddFile(file2);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var documentInfo = await _sharePointClientFixture.SharePointClient
                .GetDocumentsInfoAsync("DocumentsList", _sharePointClientFixture.GetRelativePath($"DocumentsList/{folder}"));

            // Assert
            documentInfo.Length.Should().Be(0);
            documentInfo.Name.Should().Be(folder);
            documentInfo.IsFile.Should().BeFalse();
            documentInfo.ContentBase64.Should().BeNull();
            documentInfo.Documents.Count().Should().Be(2);
            documentInfo.Extension.Should().BeNullOrEmpty();
            documentInfo.Documents.Should().Contain(d => d.Id.Equals(file1.Id));
            documentInfo.Documents.Should().Contain(d => d.Name.Equals("SubFolder"));
            documentInfo.Documents.Should().Contain(d => d.Name.Equals(file1.FileName));

            var subFolder = documentInfo.Documents.First(d => d.Name.Equals("SubFolder"));
            subFolder.Length.Should().Be(0);
            subFolder.IsFile.Should().BeFalse();
            subFolder.ContentBase64.Should().BeNull();
            subFolder.Documents.Count().Should().Be(1);
            subFolder.Extension.Should().BeNullOrEmpty();
            subFolder.Documents.Should().Contain(d => d.Id.Equals(file2.Id));
            subFolder.Documents.Should().Contain(d => d.Name.Equals(file2.FileName));
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "Getting all documents information on specific folder the SharePoint document library from absolute path")]
        public async Task SharePointClient_GetDocumentsInfoAsync_GettingAllDocumentsFromAbsolutePath()
        {
            // Arrange
            var folder = Guid.NewGuid().ToString();
            var file1 = _sharePointClientFixture.GenerateSharePointFile(folder);
            var file2 = _sharePointClientFixture.GenerateSharePointFile($"{folder}/SubFolder");

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file1);
            _sharePointClientFixture.SharePointClient.AddFile(file2);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var documentInfo = await _sharePointClientFixture.SharePointClient
                .GetDocumentsInfoAsync("DocumentsList", _sharePointClientFixture.GetAbsolutePath($"DocumentsList/{folder}"));

            // Assert
            documentInfo.Length.Should().Be(0);
            documentInfo.Name.Should().Be(folder);
            documentInfo.IsFile.Should().BeFalse();
            documentInfo.ContentBase64.Should().BeNull();
            documentInfo.Documents.Count().Should().Be(2);
            documentInfo.Extension.Should().BeNullOrEmpty();
            documentInfo.Documents.Should().Contain(d => d.Id.Equals(file1.Id));
            documentInfo.Documents.Should().Contain(d => d.Name.Equals("SubFolder"));
            documentInfo.Documents.Should().Contain(d => d.Name.Equals(file1.FileName));

            var subFolder = documentInfo.Documents.First(d => d.Name.Equals("SubFolder"));
            subFolder.Length.Should().Be(0);
            subFolder.IsFile.Should().BeFalse();
            subFolder.ContentBase64.Should().BeNull();
            subFolder.Documents.Count().Should().Be(1);
            subFolder.Extension.Should().BeNullOrEmpty();
            subFolder.Documents.Should().Contain(d => d.Id.Equals(file2.Id));
            subFolder.Documents.Should().Contain(d => d.Name.Equals(file2.FileName));
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "Getting file information the SharePoint document library from relative path")]
        public async Task SharePointClient_GetDocumentsInfoAsync_GettingFileFromRelativePath()
        {
            // Arrange
            var folder = Guid.NewGuid().ToString();
            var file = _sharePointClientFixture.GenerateSharePointFile(folder);

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var documentInfo = await _sharePointClientFixture.SharePointClient
                .GetDocumentsInfoAsync("DocumentsList", _sharePointClientFixture.GetRelativePath($"DocumentsList/{folder}/{file.FileName}"));

            // Assert
            documentInfo.Id.Should().Be(file.Id);
            documentInfo.IsFile.Should().BeTrue();
            documentInfo.Extension.Should().Be(".txt");
            documentInfo.Name.Should().Be(file.FileName);
            documentInfo.Documents.Count().Should().Be(0);
            documentInfo.Length.Should().BeGreaterThan(0);
            documentInfo.ContentBase64.Should().NotBeNullOrEmpty();
        }

        [Trait("Category", "SharePointOnline - Files")]
        [Fact(DisplayName = "Getting file information the SharePoint document library from absolute path")]
        public async Task SharePointClient_GetDocumentsInfoAsync_GettingFileFromAbsolutePath()
        {
            // Arrange
            var folder = Guid.NewGuid().ToString();
            var file = _sharePointClientFixture.GenerateSharePointFile(folder);

            // Act
            _sharePointClientFixture.SharePointClient.AddFile(file);

            await _sharePointClientFixture.SharePointClient.SaveChangesAsync();

            var documentInfo = await _sharePointClientFixture.SharePointClient
                .GetDocumentsInfoAsync("DocumentsList", _sharePointClientFixture.GetAbsolutePath($"DocumentsList/{folder}/{file.FileName}"));

            // Assert
            documentInfo.Id.Should().Be(file.Id);
            documentInfo.IsFile.Should().BeTrue();
            documentInfo.Extension.Should().Be(".txt");
            documentInfo.Name.Should().Be(file.FileName);
            documentInfo.Documents.Count().Should().Be(0);
            documentInfo.Length.Should().BeGreaterThan(0);
            documentInfo.ContentBase64.Should().NotBeNullOrEmpty();
        }
    }
}