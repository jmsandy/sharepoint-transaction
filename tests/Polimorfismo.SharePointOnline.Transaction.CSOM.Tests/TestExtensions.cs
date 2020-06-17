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

namespace Polimorfismo.SharePointOnline.Transaction.Tests
{
    /// <summary>
    /// Tests extensions.
    /// Inspired by SpecUnit's SpecificationExtensions
    /// </summary>
    /// <see cref="http://code.google.com/p/specunit-net/source/browse/trunk/src/SpecUnit/SpecificationExtensions.cs"/>
    /// <Author>Jose Mauro da Silva Sandy</Author>
    /// <Date>2020-06-07 10:00:34 PM</Date>
    public static class TestExtensions
    {
        public static void ShoudBeNull(this object actual)
        {
            Assert.Null(actual);
        }

        public static void ShouldBeTrue(this bool b)
        {
            Assert.True(b);
        }

        public static void ShouldBeFalse(this bool b)
        {
            Assert.False(b);
        }

        public static void ShouldEqual(this object actual, object expected)
        {
            Assert.Equal(expected, actual);
        }
    }
}
