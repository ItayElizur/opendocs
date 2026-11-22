using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;

namespace OfficeAi.Shared
{
    // Loads the add-in's own copy of a DLL, whatever version was asked for, when .NET cannot bind it.
    // The binding redirects in <AddIn>.dll.config do the same job, but on some machines Office never applies
    // that file. See DependencyResolver.cs.md.
    public static class DependencyResolver
    {
        private static int _registered;

        // Call first, from the add-in's static constructor. Safe to call more than once.
        public static void Register()
        {
            if (Interlocked.Exchange(ref _registered, 1) == 1) return;
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            AppDomain.CurrentDomain.AssemblyResolve += (sender, e) => Resolve(dir, e.Name);
        }

        public static Assembly Resolve(string dir, string requestedName)
        {
            string name = new AssemblyName(requestedName).Name;
            Assembly loaded = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(a => a.GetName().Name == name);
            if (loaded != null) return loaded;
            string path = Path.Combine(dir, name + ".dll");
            return File.Exists(path) ? Assembly.LoadFrom(path) : null;
        }
    }
}
