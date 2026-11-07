using System;
using System.Collections.Generic;
using System.IO;
using OfficeAi.Shared;
using Xunit;

namespace OfficeAi.Shared.Tests
{
    public class SecretStoreTests : IDisposable
    {
        private readonly string _app = "OpenDocsTests-" + Guid.NewGuid().ToString("N");

        public void Dispose()
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), _app);
            if (Directory.Exists(dir)) Directory.Delete(dir, true);
        }

        [Fact]
        public void NothingStoredLoadsAsEmpty()
        {
            Assert.Empty(SecretStore.Load(_app));
        }

        [Fact]
        public void KeysSurviveARoundTrip()
        {
            SecretStore.Save(_app, new Dictionary<string, string> { { "openai", "sk-test-123" }, { "custom", "k2" } });
            var loaded = SecretStore.Load(_app);
            Assert.Equal("sk-test-123", loaded["openai"]);
            Assert.Equal("k2", loaded["custom"]);
        }

        [Fact]
        public void TheFileDoesNotContainThePlainKey()
        {
            SecretStore.Save(_app, new Dictionary<string, string> { { "openai", "sk-plain-marker-987" } });
            string file = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), _app, "secrets.bin");
            Assert.DoesNotContain("sk-plain-marker-987", System.Text.Encoding.UTF8.GetString(File.ReadAllBytes(file)));
        }

        [Fact]
        public void ACorruptFileLoadsAsEmptyInsteadOfThrowing()
        {
            SecretStore.Save(_app, new Dictionary<string, string> { { "a", "b" } });
            string file = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), _app, "secrets.bin");
            File.WriteAllBytes(file, new byte[] { 1, 2, 3 });
            Assert.Empty(SecretStore.Load(_app));
        }
    }
}
