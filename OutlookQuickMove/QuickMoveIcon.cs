using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace OutlookQuickMove
{
    internal static class QuickMoveIcon
    {
        private const string ResourceName = "OutlookQuickMove.Assets.QuickMove.ico";

        public static void ApplyTo(Form form)
        {
            if (form == null)
            {
                return;
            }

            try
            {
                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName))
                {
                    if (stream == null)
                    {
                        QuickMoveLog.Write("Icon resource was not found.");
                        return;
                    }

                    using (var icon = new Icon(stream))
                    {
                        var formIcon = (Icon)icon.Clone();
                        form.Icon = formIcon;

                        // A Form does not dispose an externally assigned Icon, so release the
                        // cloned GDI handle when the form is disposed to avoid leaking it.
                        form.Disposed += delegate { formIcon.Dispose(); };
                    }
                }
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to load icon.", ex);
            }
        }
    }
}
