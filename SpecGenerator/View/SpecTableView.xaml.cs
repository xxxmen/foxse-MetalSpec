using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using MColor = System.Windows.Media.Color;
using DColor = System.Drawing.Color;
using MetalSpec.ViewModel;
using System.Collections.ObjectModel;

namespace MetalSpec
{
    /// <summary>
    /// Interaction logic for SpecTableView.xaml
    /// </summary>
    /// 
    public class ColorToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            DColor col = (DColor)value;
            MColor mcol = MColor.FromArgb(col.A, col.R, col.G, col.B);
            return new SolidColorBrush(mcol);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }

    }

    public partial class SpecTableView : Window
    {
        public SpecTableView()
        {
            InitializeComponent();
            //SpecTableViewModel svm = new SpecTableViewModel();
            //this.DataContext = svm;
            // specHeadDataGrid.ItemsSource = new ObservableCollection<SpecTableViewModel>() { DataContext as SpecTableViewModel };
        }



        public static IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj != null)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(depObj, i);
                    if (child != null && child is T)
                    {
                        yield return (T)child;
                    }

                    foreach (T childOfChild in FindVisualChildren<T>(child))
                    {
                        yield return childOfChild;
                    }
                }
            }
        }

        private bool CanExit()
        {
            var viewModel = specHeadDataGrid.DataContext as SpecTableViewModel;
            if (viewModel.Documents.Count > 0)
                switch (System.Windows.Forms.MessageBox.Show("Сохранить изменения в текущем документе?", "Выход из программы", System.Windows.Forms.MessageBoxButtons.YesNoCancel, System.Windows.Forms.MessageBoxIcon.Question))
                {
                    case System.Windows.Forms.DialogResult.Yes:
                        if (viewModel.SpecFilePatch != null && viewModel.SpecFilePatch != $"{Environment.GetEnvironmentVariable("USERPROFILE")}\\backup.json")
                            viewModel.SaveSpecification();
                        else
                            viewModel.SaveAsSpecification();
                        break;
                    case System.Windows.Forms.DialogResult.No:

                        break;
                    case System.Windows.Forms.DialogResult.Cancel:
                        return false;
                }
            return true;
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            string input = e.Text;

            if (input == "." || input == ",")
            {
                TextBox tb = ((TextBox)sender);
                if (tb.Text.Contains(".") || tb.Text.Contains(","))
                {
                    int caretIndex = tb.CaretIndex;
                    int charIndex = tb.Text.IndexOf('.');
                    if (charIndex < caretIndex)
                    {
                        caretIndex--;
                        if (caretIndex < 0)
                            caretIndex = 0;
                    }
                    tb.Text = tb.Text.Replace(".", "").Replace(",", "").Insert(caretIndex, input.Replace(",", "."));

                    tb.CaretIndex = caretIndex + 1;
                    e.Handled = true;
                }
                else
                {
                    int caretIndex = tb.CaretIndex;
                    tb.Text = tb.Text.Insert(caretIndex, input.Replace(",", "."));
                    if (tb.Text == "0.")
                        tb.CaretIndex = caretIndex + 2;
                    else
                        tb.CaretIndex = caretIndex + 1;
                    e.Handled = true;
                }
            }
            else if (e.Text == "0")
            {
                e.Handled = false;
            }
            else
                e.Handled = !IsTextAllowed(input);
        }

        static Regex regex = new Regex(@"[0-9]"); //regex that matches allowed text
        private static bool IsTextAllowed(string text)
        {
            bool alow = regex.IsMatch(text);
            return alow;
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

            TextBox txb = ((TextBox)sender);

            if ((string)txb.Tag == txb.Text || txb.Text == "." || txb.Text == ",")
            {
                if (txb.Text == "." || txb.Text == ",")
                {
                    txb.Text = '0' + txb.Text.Replace(",", ".");
                    txb.CaretIndex = txb.Text.Length;
                }
                return;
            }
            if (txb.Text.Length > 0)
            {
                if (txb.Text[0] == '.')
                {
                    txb.Text = '0' + txb.Text.Replace(",", "."); ;
                    return;
                }
                if (txb.Text[txb.Text.Length - 1] == '.' || txb.Text[txb.Text.Length - 1] == ',' || (txb.Text.Contains(".") && txb.Text[txb.Text.Length - 1] == '0') || (txb.Text.Contains(".") && txb.Text[txb.Text.Length - 1] == '0')) return;
            }

            if (txb.Text == "")
            {
                txb.Tag = txb.Text;
                txb.Text = "0";
            }

            txb.GetBindingExpression(TextBox.TextProperty).UpdateSource();

            //foreach (TextBlock tb in FindVisualChildren<TextBlock>(this))
            //{
            //    BindingExpression be = tb.GetBindingExpression(TextBlock.TextProperty);
            //    if (be != null)
            //    {
            //        be.UpdateTarget();
            //    }
            //}

            //foreach (ListBox lb in FindVisualChildren<ListBox>(this))
            //{
            //    BindingExpression be = lb.GetBindingExpression(ListBox.ItemsSourceProperty);
            //    if (be != null)
            //    {
            //        be.UpdateTarget();
            //    }
            //}

            if (txb.Text == "0")
            {
                txb.Text = "";
            }
            else
            {
                txb.Tag = null;
            }

        }

        private void TextBox_MouseUp(object sender, MouseButtonEventArgs e)
        {
            ((TextBox)sender).SelectAll();
        }

        private void specListBox_LayoutUpdated(object sender, System.EventArgs e)
        {
            //foreach (TextBlock tb in FindVisualChildren<TextBlock>(this))
            //{
            //    BindingExpression be = tb.GetBindingExpression(TextBlock.TextProperty);
            //    if (be != null)
            //    {
            //        be.UpdateTarget();
            //    }
            //}
        }

        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {

        }

        private void TextBox_TextChanged_1(object sender, TextChangedEventArgs e)
        {
            var textEdit = sender as TextBox;

            if (textEdit != null)
            {
                // if (!textEdit.IsKeyboardFocusWithin)
                {
                    var viewModel = specHeadDataGrid.DataContext as SpecTableViewModel;
                    if (!viewModel.IsThisNameExist(textEdit.Text))
                    {
                        // int o = -1;
                        //if (int.TryParse(textEdit.Tag, out o))
                        viewModel.RenameConstriction(textEdit.Text, (int)textEdit.Tag);
                    }
                }
            }
        }

        private void mainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!CanExit())
                e.Cancel = true;
            else
            {
                var viewModel = specHeadDataGrid.DataContext as SpecTableViewModel;
                viewModel.settings.Save();
            }
        }

        private void MenuItem_Click_1(object sender, RoutedEventArgs e)
        {
            AboutBox1 about = new AboutBox1();
            about.Show();
        }
    }
}
