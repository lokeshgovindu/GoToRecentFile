using System.Runtime.InteropServices;
using System.Windows;
using Microsoft.VisualStudio.Shell;

namespace GoToRecentFile.Options
{
    [ComVisible(true)]
    [Guid("A3F2C8E1-7B4D-4E9A-B6D5-1C8E3F7A2B9D")]
    internal class SettingsPage : UIElementDialogPage
    {
        private SettingsPageControl _control;

        protected override UIElement Child
        {
            get { return _control ?? (_control = new SettingsPageControl()); }
        }

        protected override void OnApply(PageApplyEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (e.ApplyBehavior == ApplyKind.Apply)
            {
                _control?.SaveSettings();
            }
            base.OnApply(e);
        }

        protected override void OnActivate(System.ComponentModel.CancelEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            base.OnActivate(e);
            _control?.LoadSettings();
        }
    }
}
