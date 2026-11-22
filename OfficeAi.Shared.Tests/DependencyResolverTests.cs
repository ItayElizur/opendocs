using System;
using OfficeAi.Shared;
using Xunit;

public class DependencyResolverTests
{
    private static readonly string Dir = AppDomain.CurrentDomain.BaseDirectory;

    [Fact]
    public void Resolve_OlderVersionThanShipped_ReturnsTheShippedCopy()
    {
        var asm = DependencyResolver.Resolve(Dir, "System.Runtime.CompilerServices.Unsafe, Version=4.0.4.1, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a");
        Assert.NotNull(asm);
        Assert.Equal("System.Runtime.CompilerServices.Unsafe", asm.GetName().Name);
        Assert.Equal(new Version(6, 0, 0, 0), asm.GetName().Version);
    }

    [Fact]
    public void Resolve_NotInTheFolder_ReturnsNull()
    {
        Assert.Null(DependencyResolver.Resolve(Dir, "No.Such.Assembly, Version=1.0.0.0"));
        Assert.Null(DependencyResolver.Resolve(Dir, "OfficeAi.Shared.resources, Version=1.0.0.0, Culture=he"));
    }
}
