using System;
using System.CodeDom.Compiler;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.CSharp;
using Xunit;

namespace OfficeAi.Shared.Tests
{
    public class DependencyResolverTests
    {
        [Fact]
        public void Resolve_LoadsTheFolderCopy_WhateverVersionWasAskedFor()
        {
            // A fresh assembly this process has never loaded, so Resolve has to read it from the folder.
            string name = "ResolverProbe" + Guid.NewGuid().ToString("N");
            string dir = Path.Combine(Path.GetTempPath(), name);
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, name + ".dll");
            var built = new CSharpCodeProvider().CompileAssemblyFromSource(
                new CompilerParameters { OutputAssembly = path, GenerateInMemory = false },
                "[assembly: System.Reflection.AssemblyVersion(\"6.0.0.0\")] public class Probe {}");
            Assert.False(built.Errors.HasErrors, "probe assembly did not compile");

            var asm = DependencyResolver.Resolve(dir, name + ", Version=4.0.4.1, Culture=neutral, PublicKeyToken=null");

            Assert.NotNull(asm);
            Assert.Equal(name, asm.GetName().Name);
            Assert.Equal(new Version(6, 0, 0, 0), asm.GetName().Version);
            Assert.Equal(path, asm.Location, StringComparer.OrdinalIgnoreCase);
        }

        [Fact]
        public void Resolve_NotInTheFolder_ReturnsNull()
        {
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            Assert.Null(DependencyResolver.Resolve(dir, "No.Such.Assembly, Version=1.0.0.0"));
            Assert.Null(DependencyResolver.Resolve(dir, "OfficeAi.Shared.resources, Version=1.0.0.0, Culture=he"));
        }

        // Without this line an app loads fine on most machines and fails at startup where Office ignores the .dll.config.
        [Theory]
        [InlineData("Word")]
        [InlineData("Excel")]
        [InlineData("PowerPoint")]
        [InlineData("Outlook")]
        public void EveryAddIn_RegistersTheResolver_InItsStaticConstructor(string app)
        {
            string source = File.ReadAllText(Path.Combine(TierListsAgreeTests.Repo(), app + "AiAddIn", "ThisAddIn.cs"));
            Assert.Matches(new Regex(@"static ThisAddIn\(\)\s*\{\s*DependencyResolver\.Register\(\);\s*\}"), source);
        }
    }
}
