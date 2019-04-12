using System;
using EnvDTE;

namespace BennorMcCarthy.AutoT4
{
    public class AutoT4ExtenderProvider : IExtenderProvider
    {
        public const string Name = "AutoT4ExtenderProvider";

        private readonly DTE _dte;

        public AutoT4ExtenderProvider(DTE dte)
        {
            _dte = dte ?? throw new ArgumentNullException("dte");
        }

        public object GetExtender(string extenderCatId, string extenderName, object extendeeObject, IExtenderSite extenderSite, int cookie)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            if (!CanExtend(extenderCatId, extenderName, extendeeObject))
                return null;

            if (!(extendeeObject is VSLangProj.FileProperties fileProperties))
            {
                return null;
            }
            var item = _dte.Solution.FindProjectItem(fileProperties.FullPath);
            if (item == null)
                return null;

            return new AutoT4Extender(item, extenderSite, cookie);
        }

        public bool CanExtend(string extenderCatid, string extenderName, object extendeeObject)
        {
            var fileProperties = extendeeObject as VSLangProj.FileProperties;
            return extenderName == Name &&
                   fileProperties != null &&
                   ".tt".Equals(fileProperties.Extension, StringComparison.OrdinalIgnoreCase);
        }
    }
}