using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace FlexCalc
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        DataModel Model = new DataModel();
        const string FileName = "result.xlsx";
        bool UpdatingInput;

        public MainPage()
        {
            this.InitializeComponent();

            this.NavigationCacheMode = NavigationCacheMode.Required;
        }

        private async void Page_Loaded(object sender, RoutedEventArgs e)
        {
            UpdatingInput = true;
            try
            {
                await Model.LoadSpreadsheetAsync(FileName);
                CreateCells();
                FillInput();
            }
            finally
            {
                UpdatingInput = false;
            }
        }

        private void CreateCells()
        {
            for (int i = 0; i < 10; i++)
            {
                ContentPanel.RowDefinitions.Add(new RowDefinition());
            }

            for (int r = 0; r < ContentPanel.RowDefinitions.Count; r++)
            {
                var Head = new TextBlock() { Text = "A" + (r + 1).ToString(), VerticalAlignment = VerticalAlignment.Center, FontSize = 16 };
                ContentPanel.Children.Add(Head);
                Grid.SetRow(Head, r);
                Grid.SetColumn(Head, 0);

                var Formula = new TextBox();
                ContentPanel.Children.Add(Formula);
                Grid.SetRow(Formula, r);
                Grid.SetColumn(Formula, 1);
                Formula.InputScope = new InputScope() { Names = { new InputScopeName() { NameValue = InputScopeNameValue.Formula } } };
                Formula.TextChanged += Formula_TextChanged;

                var Result = new TextBlock() { VerticalAlignment = VerticalAlignment.Center, FontSize = 16 };
                ContentPanel.Children.Add(Result);
                Grid.SetRow(Result, r);
                Grid.SetColumn(Result, 2);

            }
        }

        private void FillInput()
        {
            var boxes = from b in ContentPanel.Children where (b is TextBox) && Grid.GetColumn(b as TextBox) == 1 select b;
            foreach (TextBox b in boxes)
            {
                b.Text = Model.GetCellOrFormula(Grid.GetRow(b) + 1);
            }
        }

        async void Formula_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!Model.Loaded || UpdatingInput) return;
            TextBox tb = sender as TextBox;
            if (tb == null) return;
            Model.SetCellFromString(Grid.GetRow(tb) + 1, 1, tb.Text);
            await Model.SaveState(FileName);
            Model.Recalc();
            UpdateResults();
        }

        private void UpdateResults()
        {
            var boxes = from b in ContentPanel.Children where (b is TextBlock) && Grid.GetColumn(b as TextBlock) == 2 select b;
            foreach (TextBlock b in boxes)
            {
                b.Text = Model.GetStringFromCell(Grid.GetRow(b) + 1, 1);
            }
        }
    }
}
