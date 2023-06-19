using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace resource.package
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(CONSTANT.GUID)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.ShellInitialized_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class PreviewSTL : AsyncPackage
    {
        internal static class CONSTANT
        {
            public const string APPLICATION = "Visual Studio";
            public const string COMPANY = "Viacheslav Lozinskyi";
            public const string COPYRIGHT = "Copyright (c) 2023 by Viacheslav Lozinskyi. All rights reserved.";
            public const string DESCRIPTION = "Quick preview of STL files";
            public const string GUID = "E7B35F7F-99A5-4591-9545-34B1CB3A8FA8";
            public const string HOST = "MetaOutput";
            public const string NAME = "Preview-STL";
            public const string VERSION = "1.1.0";
            public const string PIPE = "urn:metaoutput:pipe:Preview-STL";
        }

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            extension.AnyPipe.Connect(CONSTANT.APPLICATION, CONSTANT.NAME);
            extension.AnyPipe.Register(CONSTANT.PIPE, new pipe.VSPipe());
            extension.AnyPreview.Connect(CONSTANT.APPLICATION, CONSTANT.NAME);
            extension.AnyPreview.Register(".STL", new preview.VSPreview());
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
        }

        protected override int QueryClose(out bool canClose)
        {
            extension.AnyPreview.Disconnect();
            extension.AnyPipe.Disconnect();
            canClose = true;
            return VSConstants.S_OK;
        }
    }
}
