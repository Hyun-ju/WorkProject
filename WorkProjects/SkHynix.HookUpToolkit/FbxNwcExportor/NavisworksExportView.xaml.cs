using System.Windows;

namespace SkHynix.HookUpToolkit.FbxNwcExportor {
  /// <summary>
  /// NavisworksExportView.xaml에 대한 상호 작용 논리
  /// </summary>
  public partial class NavisworksExportView : Window {
    public NavisworksExportVM VM { get { return DataContext as NavisworksExportVM; } }
    public NavisworksExportView() {
      InitializeComponent();
    }
  }
}
