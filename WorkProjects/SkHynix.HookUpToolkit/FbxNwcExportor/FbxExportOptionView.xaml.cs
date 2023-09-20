using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace SkHynix.HookUpToolkit.FbxNwcExportor {
  /// <summary>
  /// FbxExportOptionView.xaml에 대한 상호 작용 논리
  /// </summary>
  public partial class FbxExportOptionView : Window {
    public FbxExportOptionView() {
      InitializeComponent();
    }
    private void TextBox_PreviewKeyDown(object sender, KeyEventArgs e) {
      if (e.Key == Key.Space) { e.Handled = true; }
    }

    private void TextBox_OnlyNumberInput(object sender, TextCompositionEventArgs e) {
      if (int.TryParse(e.Text, out var result)) {
        return;
      } else if (e.Text.Equals(".")) {
        if (!(sender as TextBox).Text.Contains(".")) {
          return;
        }
      }

      e.Handled = true;
    }
  }
}
