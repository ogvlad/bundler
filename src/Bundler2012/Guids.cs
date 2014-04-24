// Guids.cs
// MUST match guids.h

using System;

namespace Bundler2012
{
    static class GuidList
    {
        public const string guidBundlerRunOnSavePkgString = "68CE1BF0-D4F8-4F45-BB02-F7C51995F121";
        public const string guidBundlerRunOnSaveCmdSetString = "339CEF60-F5ED-4E75-8A7A-18A153C2EB47";
        public const string guidBundlerRunOnSaveOutputWindowPane = "324112CA-3C02-451A-9E53-A9C66A31E833";

        public static readonly Guid guidBundlerRunOnSaveCmdSet = new Guid(guidBundlerRunOnSaveCmdSetString);
    };
}