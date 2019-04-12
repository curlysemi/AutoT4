using EnvDTE;
using System;
using System.Runtime.InteropServices;

namespace BennorMcCarthy.AutoT4
{
    [CLSCompliant(false)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class AutoT4Extender : AutoT4ProjectItemSettings
    {
        private readonly IExtenderSite _extenderSite;
        private readonly int _cookie;

        public AutoT4Extender(ProjectItem item, IExtenderSite extenderSite, int cookie)
            : base(item)
        {
            _extenderSite = extenderSite ?? throw new ArgumentNullException("extenderSite");
            _cookie = cookie;
        }

        ~AutoT4Extender()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
#pragma warning disable VSTHRD010
                _extenderSite?.NotifyDelete(_cookie);
#pragma warning restore VSTHRD010
            }
            catch
            {
            }
        }
    }
}