using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.IsolatedStorage;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using VirtualDesktopGridSwitcher.Settings;

namespace VirtualDesktopGridSwitcher {
    class SysTrayProcess : IDisposable {

        private string iconsDirName = "Icons";
        private NotifyIcon notifyIcon;
        private Icon[] desktopIcons;

        private ContextMenus contextMenu;

        public SysTrayProcess(SettingValues settings) {
            notifyIcon = new NotifyIcon();
            notifyIcon.Visible = true;
            notifyIcon.Text = "Virtual Desktop Grid Switcher";

            contextMenu = new ContextMenus(settings);
            notifyIcon.ContextMenuStrip = contextMenu.MenuStrip;

            LoadIconImages();
        }

        public void Dispose() {
            notifyIcon.Dispose();
        }

        public void ShowIconForDesktop(int desktopIndex) {
            notifyIcon.Icon = desktopIcons[desktopIndex];
        }

        private void LoadIconImages() {
            using (IsolatedStorageFile isoStore = IsolatedStorageFile.GetStore(IsolatedStorageScope.User | IsolatedStorageScope.Domain | IsolatedStorageScope.Assembly, null, null)) {

                if (!isoStore.DirectoryExists(iconsDirName)) {
                    isoStore.CreateDirectory(iconsDirName);
                    var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                    var iconDir = Path.Combine(baseDir, "Icons");

                    foreach (var f in Directory.GetFiles(iconsDirName, "*.ico")
                        .OrderBy(f => int.Parse(Path.GetFileNameWithoutExtension(f)))) {

                        using (var fsIn = File.OpenRead(f)) {
                            using (var fsOut = isoStore.OpenFile(Path.Combine(iconsDirName, Path.GetFileName(f)), FileMode.Create)) {
                                fsIn.CopyTo(fsOut);
                            }
                        }
                    }
                }

                var icons = new List<Icon>();
                foreach (var f in isoStore.GetFileNames(Path.Combine(iconsDirName, "*.ico"))
                    .OrderBy(f => int.Parse(Path.GetFileNameWithoutExtension(f)))) {

                    using (var fs = isoStore.OpenFile(Path.Combine(iconsDirName, f), FileMode.Open)) {
                        icons.Add(new Icon(fs));
                    }
                }

                desktopIcons = icons.ToArray();
            }
        }

    }
}
