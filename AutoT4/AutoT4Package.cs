using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using EnvDTE;
using System.Linq;
using System.Collections.Generic;
using System.Threading;
using System;
using Microsoft;
using Task = System.Threading.Tasks.Task;


namespace BennorMcCarthy.AutoT4
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidAutoT4PkgString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExistsAndFullyLoaded_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideOptionPage(typeof(Options), Options.CategoryName, Options.PageName, 1000, 1001, false)]

    public sealed class AutoT4Package : AsyncPackage
    {
        private DTE _dte;
        private BuildEvents _buildEvents;
        private ObjectExtenders _objectExtenders;
        private AutoT4ExtenderProvider _extenderProvider;
        private readonly List<int> _extenderProviderCookies = new List<int>();

        private Options Options
        {
            get { return (Options)GetDialogPage(typeof(Options)); }
        }

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await base.InitializeAsync(cancellationToken, progress);
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            _dte = await GetServiceAsync(typeof(DTE)) as DTE;
            Assumes.Present(_dte);

            await RegisterExtenderProviderAsync(VSConstants.CATID.CSharpFileProperties_string);
            await RegisterExtenderProviderAsync(VSConstants.CATID.VBFileProperties_string);

            await RegisterEventsAsync();
        }

        private async Task RegisterEventsAsync()
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
            _buildEvents = _dte.Events.BuildEvents;
            _buildEvents.OnBuildBegin += OnBuildBegin;
            _buildEvents.OnBuildDone += OnBuildDone;
        }

        private async Task RegisterExtenderProviderAsync(string catId)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
            const string name = AutoT4ExtenderProvider.Name;

            _objectExtenders = _objectExtenders ?? await GetServiceAsync(typeof(ObjectExtenders)) as ObjectExtenders;
            Assumes.Present(_objectExtenders);

            _extenderProvider = _extenderProvider ?? new AutoT4ExtenderProvider(_dte);
            _extenderProviderCookies.Add(_objectExtenders.RegisterExtenderProvider(catId, name, _extenderProvider));
        }

        private void OnBuildBegin(vsBuildScope scope, vsBuildAction action)
        {
            RunTemplates(scope, RunOnBuild.BeforeBuild, Options.RunOnBuild == DefaultRunOnBuild.BeforeBuild);
        }

        private void OnBuildDone(vsBuildScope scope, vsBuildAction action)
        {
            RunTemplates(scope, RunOnBuild.AfterBuild, Options.RunOnBuild == DefaultRunOnBuild.AfterBuild);
        }

        private void RunTemplates(vsBuildScope scope, RunOnBuild buildEvent, bool runIfDefault)
        {
            _dte.GetProjectsWithinBuildScope(scope)
                .FindT4ProjectItems()
                .ThatShouldRunOn(buildEvent, runIfDefault)
                .ToList()
                .ForEach(item => item.RunTemplate());
        }

        protected override int QueryClose(out bool canClose)
        {
            int result = base.QueryClose(out canClose);
            if (!canClose)
                return result;

            if (_buildEvents != null)
            {
                _buildEvents.OnBuildBegin -= OnBuildBegin;
                _buildEvents.OnBuildDone -= OnBuildDone;
                _buildEvents = null;
            }
            return result;
        }
    }
}
