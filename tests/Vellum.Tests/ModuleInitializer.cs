using System.Runtime.CompilerServices;
using VerifyTests;

namespace Vellum.Tests;

public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Initialize()
    {
        VerifyOpenXml.Initialize();

        // Auto-verify docx binary changes (timestamps differ between runs)
        // Text content comparison via #00.txt/#01.txt files provides the actual test validation
        VerifierSettings.AutoVerify(includeBuildServer: false);
    }
}
