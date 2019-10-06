using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace CustomReportsManager {
	/// <summary>
	/// Interaction logic for WindowSqlQueryView.xaml
	/// </summary>
	public partial class WindowSqlQueryView : Window {
		public WindowSqlQueryView(CustomReports.ItemReport itemReport) {
			InitializeComponent();

			if (itemReport == null) {
				MessageBox.Show(this, "ItemReport is null", "", MessageBoxButton.OK, MessageBoxImage.Error);
				return;
			}

			TextBoxQuery.Text = itemReport.Query;
			Title = "Редактирование запроса для: '" + itemReport.Name + "'";
			TextBlockAbout.Text = "Формируемый отчет будет выгружать данные из БД в соответствии с запросом, без дополнительной обработки." + Environment.NewLine +
				"Для удобства пользователей отчета называйте поля отчета понятными именами, они будут использоваться как заголовки." + Environment.NewLine +
				Environment.NewLine + "В тексте запроса можно использовать два параметра - @dateBegin и @dateEnd" + Environment.NewLine +
				"Вместо этих параметров будут подставляться даты, за которые необходимо выгрузить данные";

			Closed += (s, e) => {
				string queryEntered = TextBoxQuery.Text;
				itemReport.Query = queryEntered;
			};

			Loaded += (s, e) => {
				using (var stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("CustomReportsManager.sql.xshd")) {
					using (var reader = new System.Xml.XmlTextReader(stream)) {
						TextBoxQuery.SyntaxHighlighting =
							ICSharpCode.AvalonEdit.Highlighting.Xshd.HighlightingLoader.Load(reader,
							ICSharpCode.AvalonEdit.Highlighting.HighlightingManager.Instance);
					}
				}

				TextBoxQuery.Focus();
			};
		}

		private void ButtonClose_Click(object sender, RoutedEventArgs e) {
			Close();
		}
	}
}
