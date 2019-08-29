using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

namespace FlexCalc
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        DataModel Model = new DataModel();
        readonly string FileName = Path.Combine(Windows.Storage.ApplicationData.Current.TemporaryFolder.Path,  "result.xlsx");
        bool UpdatingInput;

        public MainPage()
        {
            this.InitializeComponent();
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            UpdatingInput = true;
            try
            {
                Model.LoadSpreadsheet(FileName);
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

                var Formula = new TextBox() { TextWrapping = TextWrapping.Wrap };
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

        void Formula_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!Model.Loaded || UpdatingInput) return;
            TextBox tb = sender as TextBox;
            if (tb == null) return;
            Model.SetCellFromString(Grid.GetRow(tb) + 1, 1, tb.Text);
            Model.SaveState(FileName);
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
