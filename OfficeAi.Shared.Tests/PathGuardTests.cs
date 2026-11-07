using System.IO;
using OfficeAi.Shared;
using Xunit;

namespace OfficeAi.Shared.Tests
{
    public class PathGuardTests
    {
        [Theory]
        [InlineData("\\\\attacker\\share\\a.png")]
        [InlineData("//attacker/share/a.png")]
        [InlineData("\\\\?\\UNC\\attacker\\share\\a.png")]
        [InlineData("pictures\\a.png")]
        [InlineData("a.png")]
        public void NetworkAndRelativePathsAreRefused(string path)
        {
            Assert.Contains("local file path", PathGuard.LocalFileError(path));
        }

        [Theory]
        [InlineData("http://example.com/a.png")]
        [InlineData("HTTPS://example.com/a.png")]
        [InlineData("file://host/a.png")]
        public void RemoteUrlsAreRefused(string path)
        {
            Assert.Contains("remote URLs", PathGuard.LocalFileError(path));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void BlankIsRefused(string path)
        {
            Assert.Contains("required", PathGuard.LocalFileError(path));
        }

        [Fact]
        public void AMissingFileOnADriveIsReportedAsMissing()
        {
            Assert.Contains("no file at", PathGuard.LocalFileError("C:\\definitely\\not\\here\\a.png"));
        }

        [Fact]
        public void AnExistingLocalFileIsAccepted()
        {
            string file = Path.GetTempFileName();
            try { Assert.Null(PathGuard.LocalFileError(file)); }
            finally { File.Delete(file); }
        }
    }
}
