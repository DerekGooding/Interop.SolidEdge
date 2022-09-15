using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace GenerateXmlDocumentation
{
    class Program
    {
        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode)]
        static extern int LoadTypeLib(string szFile, out ITypeLib typeLib);

        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        static extern ITypeLib LoadRegTypeLib(ref Guid rguid, int wVerMajor, int wVerMinor, int lcid);

        static FileInfo? _targetPath;
        static DirectoryInfo? _targetDir;

        static int Main(string[] args)
        {
            try
            {
                // args[0] should be $(TargetPath)
                if (args.Length == 1)
                {
                    _targetPath = new FileInfo(args[0]);
                    Console.WriteLine(_targetPath);
                    if (_targetPath.Exists)
                    {
                        _targetDir = _targetPath.Directory;
                        if(_targetDir != null)
                            foreach (FileInfo fileInfo in _targetDir.EnumerateFiles("Interop.*.dll", SearchOption.TopDirectoryOnly))
                            {
                                Console.WriteLine("Building documentation for {0}.", fileInfo.FullName);
                                if (fileInfo.Name.Equals(_targetPath.Name, StringComparison.OrdinalIgnoreCase))
                                {
                                    // Skip Interop.SolidEdge.dll.
                                    continue;
                                }

                                var interopAssembly = Assembly.LoadFrom(fileInfo.FullName);
                                GenerateDocumentation(interopAssembly);
                            }

                        return 0;
                    }
                    else
                    {
                        throw new Exception(String.Format("$(TargetPath) {0} does not exist", args[0]));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return -1;
        }

        static void GenerateDocumentation(Assembly assembly)
        {
            var dictionary = GetHelpStrings(assembly);
            XmlDocumentationHelper.BuildForAssembly(assembly, dictionary);
        }

        public static Dictionary<string, string> GetHelpStrings(Assembly assembly)
        {
            Dictionary<string, string> dictionary = new();

            var a = assembly.CustomAttributes.FirstOrDefault(x => x.AttributeType.Equals(typeof(ImportedFromTypeLibAttribute)));
            var b = assembly.CustomAttributes.FirstOrDefault(x => x.AttributeType.Equals(typeof(GuidAttribute)));
            var c = assembly.CustomAttributes.FirstOrDefault(x => x.AttributeType.Equals(typeof(TypeLibVersionAttribute)));

            if (a != null)
            {
                Guid guid = Guid.Parse(String.Format("{0}", b.ConstructorArguments[0].Value));

                    int? wVerMajor = (int)c.ConstructorArguments[0].Value;
                    int? wVerMinor = (int)c.ConstructorArguments[1].Value;

                    ITypeLib? typeLib = null;
                    typeLib = LoadRegTypeLib(ref guid, (int)wVerMajor, (int)wVerMinor, 0);

                    typeLib.GetDocumentation(-1, out string strLibName, out string strLibDocString, out int dwLibHelpContext, out string strLibHelpFile);

                    int count = typeLib.GetTypeInfoCount();

                    // Loop through types.
                    for (int i = 0; i < count; i++)
                    {
                        typeLib.GetTypeInfo(i, out ITypeInfo typeInfo);

                        IntPtr pTypeAttr = IntPtr.Zero;

                        typeInfo.GetTypeAttr(out pTypeAttr);
                        TYPEATTR typeAttr = (TYPEATTR)Marshal.PtrToStructure(pTypeAttr, typeof(TYPEATTR));

                        // Skip type if it is hidden.
                        if (typeAttr.wTypeFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FHIDDEN) == true)
                        {
                            continue;
                        }

                        typeInfo.GetDocumentation(-1, out string? strTypeName, out string? strTypeDocString, out int dwTypeHelpContext, out string? strTypeHelpFile);

                        string typeKey = String.Format("{0}.{1}", strLibName, strTypeName);
                        dictionary.Add(typeKey, strTypeDocString);

                        for (int j = 0; j < typeAttr.cFuncs; j++)
                        {
                            IntPtr pFuncDesc = IntPtr.Zero;
                            typeInfo.GetFuncDesc(j, out pFuncDesc);

                            FUNCDESC funcDesc = (FUNCDESC)Marshal.PtrToStructure(pFuncDesc, typeof(FUNCDESC));

                            typeInfo.GetDocumentation(funcDesc.memid, out string? strMemberName, out string? strMemberDocString, out int dwMemberHelpContext, out string? strMemberHelpFile);

                            string memberKey = string.Format("{0}.{1}", typeKey, strMemberName);

                            if (!dictionary.ContainsKey(memberKey))
                            {
                                dictionary.Add(memberKey, strMemberDocString);
                            }

                            typeInfo.ReleaseFuncDesc(pFuncDesc);
                        }

                        typeInfo.ReleaseTypeAttr(pTypeAttr);

                    }
                
            }

            return dictionary;
        }
    }
}
